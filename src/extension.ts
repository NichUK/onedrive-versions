import * as path from "node:path";
import { execSync } from "node:child_process";
import { PublicClientApplication } from "@azure/msal-node";
import * as vscode from "vscode";

const CONTENT_SCHEME = "onedrive-version";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

interface GraphVersion {
  id: string;
  lastModifiedDateTime: string;
  size?: number;
  lastModifiedBy?: {
    user?: {
      displayName?: string;
    };
  };
}

interface GraphDriveItem {
  id: string;
  name: string;
  parentReference?: {
    driveId?: string;
  };
}

interface GraphDrive {
  id: string;
  name?: string;
  driveType?: string;
  webUrl?: string;
}

interface VersionContext {
  driveId: string;
  itemId: string;
  versions: GraphVersion[];
  selectedIndex: number;
}

interface Mapping {
  localRoot: string;
  driveId?: string;
  remoteRoot?: string;
  urlNamespace?: string;
  fullRemotePath?: string;
}

interface RequestOptions {
  interactive?: boolean;
}

class OneDriveClient {
  private readonly contextCache = new Map<string, VersionContext>();
  private msalApp?: PublicClientApplication;
  private msalAccountHomeId?: string;
  private msalClientId?: string;
  private msalTenantId?: string;
  private readonly windowsRegistryMappings: Mapping[] = this.getMappingsFromWindowsRegistry();

  public getAuthMode(): "vscode" | "deviceCode" {
    const cfg = vscode.workspace.getConfiguration("onedriveVersions");
    return cfg.get<"vscode" | "deviceCode">("auth.mode", "vscode");
  }

  public hasDeviceCodeClientId(): boolean {
    const cfg = vscode.workspace.getConfiguration("onedriveVersions");
    const clientId = cfg.get<string>("auth.clientId", "").trim();
    return clientId.length > 0;
  }

  public async connectAccount(): Promise<void> {
    await this.getAccessToken();
  }

  public async setAuthMode(mode: "vscode" | "deviceCode"): Promise<void> {
    const cfg = vscode.workspace.getConfiguration("onedriveVersions");
    await cfg.update("auth.mode", mode, vscode.ConfigurationTarget.Global);
  }

  public async loadVersionsForFile(localPath: string, options?: RequestOptions): Promise<VersionContext> {
    const resolved = path.resolve(localPath);
    const mapping = this.resolveBestMapping(resolved);
    if (!mapping) {
      throw new Error("File is not inside a detected OneDrive root.");
    }

    const relativeSegments = this.toRelativeSegments(mapping, resolved);
    const remotePath = this.toRemotePath(mapping, relativeSegments);
    let item: GraphDriveItem | undefined;
    try {
      item = await this.getDriveItem(mapping, remotePath, options);
    } catch (error) {
      if (!isGraphNotFound(error) && !isGraphAccessDenied(error)) {
        throw error;
      }

      try {
        item = await this.getDriveItemByDriveWebUrl(mapping, relativeSegments, options);
      } catch (driveUrlError) {
        if (!isGraphNotFound(driveUrlError) && !isGraphAccessDenied(driveUrlError)) {
          throw driveUrlError;
        }
      }

      if (!item) {
        const shareRoots = [mapping.fullRemotePath, mapping.urlNamespace]
          .filter((value): value is string => Boolean(value && value.trim().length > 0))
          .map((value) => normalizeShareBaseUrl(value));
        if (!shareRoots.length) {
          throw error;
        }
        item = await this.getDriveItemFromShareRoots(shareRoots, relativeSegments, options);
      }
    }

    if (!item) {
      throw new Error("itemNotFound: unable to resolve remote item for local OneDrive path.");
    }

    const driveId = item.parentReference?.driveId ?? mapping.driveId;

    if (!driveId) {
      throw new Error("Unable to determine driveId for this file.");
    }

    const versions = await this.getVersions(driveId, item.id, options);
    const sorted = [...versions].sort((a, b) => {
      return new Date(b.lastModifiedDateTime).getTime() - new Date(a.lastModifiedDateTime).getTime();
    });

    if (!sorted.length) {
      throw new Error("No OneDrive versions were returned for this file.");
    }

    const versionContext: VersionContext = {
      driveId,
      itemId: item.id,
      versions: sorted,
      selectedIndex: 0
    };

    this.contextCache.set(resolved, versionContext);
    return versionContext;
  }

  public getCachedContext(localPath: string): VersionContext | undefined {
    return this.contextCache.get(path.resolve(localPath));
  }

  public clearCachedContext(localPath: string): void {
    this.contextCache.delete(path.resolve(localPath));
  }

  public async downloadVersionBytes(localPath: string, versionId: string): Promise<Uint8Array> {
    const context = this.getCachedContext(localPath) ?? (await this.loadVersionsForFile(localPath));
    const endpoint = `${GRAPH_BASE}/drives/${encodeURIComponent(context.driveId)}/items/${encodeURIComponent(context.itemId)}/versions/${encodeURIComponent(versionId)}/content`;
    return this.fetchBinary(endpoint);
  }

  public findOneDriveRoot(localPath: string): Mapping | undefined {
    return this.resolveBestMapping(path.resolve(localPath));
  }

  private resolveBestMapping(localPath: string): Mapping | undefined {
    const configured = this.getMappingsFromConfig();
    const envMappings = this.getMappingsFromEnvironment();
    const registryMappings = this.windowsRegistryMappings;
    const inferred = this.inferMappingFromPath(localPath);
    const baseCandidates = [...configured, ...envMappings, ...registryMappings];
    const candidates = inferred ? [...baseCandidates, inferred] : baseCandidates;

    const matches = candidates
      .map((m) => ({ m, root: normalizeLocalRoot(m.localRoot) }))
      .filter(({ root }) => isPathWithin(localPath, root))
      .sort((a, b) => b.root.length - a.root.length);

    if (!matches.length) {
      return undefined;
    }

    return {
      ...matches[0].m,
      localRoot: matches[0].root
    };
  }

  private inferMappingFromPath(localPath: string): Mapping | undefined {
    const parsed = path.parse(localPath);
    const segments = localPath
      .slice(parsed.root.length)
      .split(path.sep)
      .filter((segment) => segment.length > 0);

    const oneDriveIndex = segments.findIndex((segment) => /^onedrive(\b|[ -])/i.test(segment));
    if (oneDriveIndex < 0) {
      return undefined;
    }

    const rootSegments = segments.slice(0, oneDriveIndex + 1);
    const inferredRoot = path.join(parsed.root, ...rootSegments);
    return { localRoot: inferredRoot };
  }

  private getMappingsFromConfig(): Mapping[] {
    const cfg = vscode.workspace.getConfiguration("onedriveVersions");
    const mappings = cfg.get<Mapping[]>("mappings", []);
    return mappings.filter((m) => typeof m.localRoot === "string" && m.localRoot.trim().length > 0);
  }

  private getMappingsFromEnvironment(): Mapping[] {
    const envRoots = [process.env.OneDrive, process.env.OneDriveCommercial, process.env.OneDriveConsumer]
      .filter((value): value is string => Boolean(value && value.trim().length > 0))
      .map((localRoot) => ({ localRoot }));

    // Remove duplicates while preserving order.
    const deduped: Mapping[] = [];
    const seen = new Set<string>();
    for (const mapping of envRoots) {
      const normalized = normalizeLocalRoot(mapping.localRoot);
      if (!seen.has(normalized)) {
        seen.add(normalized);
        deduped.push({ ...mapping, localRoot: normalized });
      }
    }
    return deduped;
  }

  private getMappingsFromWindowsRegistry(): Mapping[] {
    if (process.platform !== "win32") {
      return [];
    }

    try {
      const command = "reg query \"HKCU\\Software\\SyncEngines\\Providers\\OneDrive\" /s";
      const output = execSync(command, { stdio: ["ignore", "pipe", "ignore"], encoding: "utf8" });
      const mappingsByKey = new Map<string, Mapping>();
      const seen = new Set<string>();
      let currentKey = "";
      const lines = output.split(/\r?\n/);

      for (const line of lines) {
        const keyMatch = line.match(/^HKEY_CURRENT_USER\\Software\\SyncEngines\\Providers\\OneDrive\\(.+)$/i);
        if (keyMatch?.[1]) {
          currentKey = keyMatch[1].trim();
          if (!mappingsByKey.has(currentKey)) {
            mappingsByKey.set(currentKey, { localRoot: "" });
          }
          continue;
        }

        if (!currentKey) {
          continue;
        }

        const valueMatch = line.match(/^\s*(MountPoint|UrlNamespace|FullRemotePath)\s+REG_\w+\s*(.*)\s*$/i);
        if (!valueMatch?.[1]) {
          continue;
        }

        const mapping = mappingsByKey.get(currentKey);
        if (!mapping) {
          continue;
        }

        const name = valueMatch[1];
        const rawValue = (valueMatch[2] ?? "").trim();
        if (name === "MountPoint") {
          mapping.localRoot = rawValue;
        } else if (name === "UrlNamespace") {
          mapping.urlNamespace = rawValue;
        } else if (name === "FullRemotePath") {
          mapping.fullRemotePath = rawValue;
        }
      }

      const mappings: Mapping[] = [];
      for (const mapping of mappingsByKey.values()) {
        if (!mapping.localRoot || !mapping.localRoot.trim()) {
          continue;
        }

        const localRoot = normalizeLocalRoot(mapping.localRoot);
        if (!seen.has(localRoot)) {
          seen.add(localRoot);
          mappings.push({
            localRoot,
            urlNamespace: mapping.urlNamespace?.trim() || undefined,
            fullRemotePath: mapping.fullRemotePath?.trim() || undefined
          });
        }
      }
      return mappings;
    } catch {
      return [];
    }
  }

  private toRelativeSegments(mapping: Mapping, localPath: string): string[] {
    const root = normalizeLocalRoot(mapping.localRoot);
    const relative = path.relative(root, localPath);
    if (relative.startsWith("..") || path.isAbsolute(relative)) {
      throw new Error("File is outside the mapped local OneDrive root.");
    }

    return relative.split(path.sep).filter((segment) => segment.length > 0);
  }

  private toRemotePath(mapping: Mapping, relativeSegments: string[]): string {
    const remoteRoot = normalizeRemoteRoot(mapping.remoteRoot ?? "/");
    const encodedRelativeSegments = relativeSegments.map((segment) => encodeURIComponent(segment));

    const rootSegments = remoteRoot
      .split("/")
      .filter((segment) => segment.length > 0)
      .map((segment) => encodeURIComponent(segment));

    const joined = [...rootSegments, ...encodedRelativeSegments].join("/");
    return `/${joined}`;
  }

  private async getDriveItem(mapping: Mapping, remotePath: string, options?: RequestOptions): Promise<GraphDriveItem> {
    const remotePathCandidates = buildRemotePathCandidates(remotePath);
    const driveId = mapping.driveId;

    if (driveId) {
      for (const candidatePath of remotePathCandidates) {
        const endpoint = `${GRAPH_BASE}/drives/${encodeURIComponent(driveId)}/root:${candidatePath}?$select=id,name,parentReference`;
        try {
          return await this.fetchJson<GraphDriveItem>(endpoint, options);
        } catch (error) {
          if (!isGraphNotFound(error)) {
            throw error;
          }
        }
      }
      throw new Error("itemNotFound: path was not found in configured drive mapping.");
    }

    for (const candidatePath of remotePathCandidates) {
      const myDriveEndpoint = `${GRAPH_BASE}/me/drive/root:${candidatePath}?$select=id,name,parentReference`;
      try {
        return await this.fetchJson<GraphDriveItem>(myDriveEndpoint, options);
      } catch (error) {
        if (!isGraphNotFound(error)) {
          throw error;
        }
      }
    }

    const drives = await this.fetchJson<{ value: GraphDrive[] }>(
      `${GRAPH_BASE}/me/drives?$select=id,name,driveType`,
      options
    );

    for (const drive of drives.value ?? []) {
      for (const candidatePath of remotePathCandidates) {
        const endpoint = `${GRAPH_BASE}/drives/${encodeURIComponent(drive.id)}/root:${candidatePath}?$select=id,name,parentReference`;
        try {
          return await this.fetchJson<GraphDriveItem>(endpoint, options);
        } catch (error) {
          if (!isGraphNotFound(error)) {
            throw error;
          }
        }
      }
    }

    throw new Error("itemNotFound: path was not found in /me/drive or any accessible /me/drives entries (including trimmed-path fallback).");
  }

  private async getDriveItemByDriveWebUrl(
    mapping: Mapping,
    relativeSegments: string[],
    options?: RequestOptions
  ): Promise<GraphDriveItem> {
    const shareRoots = [mapping.fullRemotePath, mapping.urlNamespace]
      .filter((value): value is string => Boolean(value && value.trim().length > 0))
      .map((value) => normalizeShareBaseUrl(value));

    if (!shareRoots.length) {
      throw new Error("itemNotFound: no registry URL metadata available.");
    }

    const targetUrls = shareRoots.map((root) => appendPathSegmentsToUrl(root, relativeSegments));
    const drives = await this.fetchJson<{ value: GraphDrive[] }>(
      `${GRAPH_BASE}/me/drives?$select=id,name,driveType,webUrl`,
      options
    );

    for (const drive of drives.value ?? []) {
      const driveWebUrl = drive.webUrl ? normalizeShareBaseUrl(drive.webUrl) : "";
      if (!driveWebUrl) {
        continue;
      }
      for (const targetUrl of targetUrls) {
        const relative = getRelativePathByUrlPrefix(targetUrl, driveWebUrl);
        if (relative === undefined) {
          continue;
        }
        const encodedRelative = relative
          .split("/")
          .filter((s) => s.length > 0)
          .map((s) => encodeURIComponent(s))
          .join("/");
        const candidatePath = encodedRelative ? `/${encodedRelative}` : "/";
        const endpoint = `${GRAPH_BASE}/drives/${encodeURIComponent(drive.id)}/root:${candidatePath}?$select=id,name,parentReference`;
        try {
          return await this.fetchJson<GraphDriveItem>(endpoint, options);
        } catch (error) {
          if (!isGraphNotFound(error) && !isGraphAccessDenied(error)) {
            throw error;
          }
        }
      }
    }

    throw new Error("itemNotFound: registry URL fallback could not map file to an accessible drive webUrl.");
  }

  private async getDriveItemFromShareRoots(
    shareRoots: string[],
    relativeSegments: string[],
    options?: RequestOptions
  ): Promise<GraphDriveItem> {
    for (const shareRoot of shareRoots) {
      const shareUrl = appendPathSegmentsToUrl(shareRoot, relativeSegments);
      const shareId = toGraphShareId(shareUrl);
      const endpoint = `${GRAPH_BASE}/shares/${shareId}/driveItem?$select=id,name,parentReference`;
      try {
        return await this.fetchJson<GraphDriveItem>(endpoint, options);
      } catch (error) {
        if (!isGraphNotFound(error)) {
          throw error;
        }
      }
    }
    throw new Error("itemNotFound: file could not be resolved via registry share URL metadata.");
  }

  private async getVersions(driveId: string, itemId: string, options?: RequestOptions): Promise<GraphVersion[]> {
    const endpoint = `${GRAPH_BASE}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/versions?$select=id,lastModifiedDateTime,size,lastModifiedBy`;
    const response = await this.fetchJson<{ value: GraphVersion[] }>(endpoint, options);
    return response.value ?? [];
  }

  private async getAccessToken(options?: RequestOptions): Promise<string> {
    const interactive = options?.interactive ?? true;
    const authMode = this.getAuthMode();
    if (authMode === "deviceCode") {
      return this.getAccessTokenViaDeviceCode({ interactive });
    }

    const scopes = ["Files.Read.All"];
    try {
      const session = await vscode.authentication.getSession("microsoft", scopes, { createIfNone: interactive });
      if (!session) {
        throw new Error("AUTH_REQUIRED");
      }
      return session.accessToken;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (message.includes("AADSTS65002")) {
        throw new Error(
          "Tenant policy blocked VS Code Microsoft auth for Graph (AADSTS65002). Use 'OneDrive: Connect Microsoft Account' and switch to device code auth."
        );
      }
      throw error;
    }
  }

  private async getAccessTokenViaDeviceCode(options?: RequestOptions): Promise<string> {
    const interactive = options?.interactive ?? true;
    const cfg = vscode.workspace.getConfiguration("onedriveVersions");
    const clientId = cfg.get<string>("auth.clientId", "").trim();
    const tenantId = cfg.get<string>("auth.tenantId", "organizations").trim() || "organizations";
    const scopes = ["https://graph.microsoft.com/Files.Read.All"];

    if (!clientId) {
      throw new Error(
        "Device code auth requires onedriveVersions.auth.clientId. Run 'OneDrive: Open Setup Guide' to configure your Entra app."
      );
    }

    if (!this.msalApp || this.msalClientId !== clientId || this.msalTenantId !== tenantId) {
      this.msalApp = new PublicClientApplication({
        auth: {
          clientId,
          authority: `https://login.microsoftonline.com/${tenantId}`
        }
      });
      this.msalClientId = clientId;
      this.msalTenantId = tenantId;
      this.msalAccountHomeId = undefined;
    }

    if (this.msalAccountHomeId) {
      const accounts = await this.msalApp.getTokenCache().getAllAccounts();
      const account = accounts.find((a) => a.homeAccountId === this.msalAccountHomeId);
      if (account) {
        try {
          const silent = await this.msalApp.acquireTokenSilent({ account, scopes });
          if (silent?.accessToken) {
            return silent.accessToken;
          }
        } catch {
          // Fall back to device code.
        }
      }
    }

    if (!interactive) {
      throw new Error("AUTH_REQUIRED");
    }

    let browserOpened = false;
    const interactiveToken = await this.msalApp.acquireTokenByDeviceCode({
      scopes,
      deviceCodeCallback: (response) => {
        if (!browserOpened && response.verificationUri) {
          browserOpened = true;
          void vscode.env.openExternal(vscode.Uri.parse(response.verificationUri));
        }

        if (response.userCode) {
          void vscode.env.clipboard.writeText(response.userCode);
          void vscode.window.showInformationMessage(
            `Device code copied to clipboard: ${response.userCode}. Complete sign-in in your browser.`
          );
          return;
        }

        void vscode.window.showInformationMessage(response.message);
      }
    });

    if (!interactiveToken?.accessToken) {
      throw new Error("Device code sign-in did not return a Graph access token.");
    }
    if (interactiveToken.account?.homeAccountId) {
      this.msalAccountHomeId = interactiveToken.account.homeAccountId;
    }
    return interactiveToken.accessToken;
  }

  private async fetchJson<T>(url: string, options?: RequestOptions): Promise<T> {
    const token = await this.getAccessToken(options);
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });
    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Graph request failed (${response.status}): ${body}`);
    }
    return (await response.json()) as T;
  }

  private async fetchBinary(url: string, options?: RequestOptions): Promise<Uint8Array> {
    const token = await this.getAccessToken(options);
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });
    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Graph content request failed (${response.status}): ${body}`);
    }
    const arrayBuffer = await response.arrayBuffer();
    return new Uint8Array(arrayBuffer);
  }
}

class OneDriveVersionContentProvider implements vscode.TextDocumentContentProvider {
  private readonly onDidChangeEmitter = new vscode.EventEmitter<vscode.Uri>();
  public readonly onDidChange = this.onDidChangeEmitter.event;

  public constructor(private readonly client: OneDriveClient) {}

  public async provideTextDocumentContent(uri: vscode.Uri): Promise<string> {
    const localPath = decodeURIComponent(uri.query.replace(/^local=/, ""));
    const versionId = decodeURIComponent(uri.fragment.replace(/^version=/, ""));

    if (!localPath || !versionId) {
      return "Invalid OneDrive version URI.";
    }

    const bytes = await this.client.downloadVersionBytes(localPath, versionId);
    return this.decodeAsText(bytes);
  }

  private decodeAsText(bytes: Uint8Array): string {
    const decoder = new TextDecoder("utf-8", { fatal: false });
    const text = decoder.decode(bytes);
    if (text.includes("\u0000")) {
      return "This version appears to be binary content. Use 'Save Version As...' or 'Restore Selected Version'.";
    }
    return text;
  }
}

export function activate(context: vscode.ExtensionContext): void {
  const client = new OneDriveClient();
  const contentProvider = new OneDriveVersionContentProvider(client);
  const onboardingKey = "onedriveVersions.onboardingPromptShown";

  context.subscriptions.push(vscode.workspace.registerTextDocumentContentProvider(CONTENT_SCHEME, contentProvider));

  const openSetupGuide = async (): Promise<void> => {
    const readmeUri = vscode.Uri.joinPath(context.extensionUri, "README.md");
    const doc = await vscode.workspace.openTextDocument(readmeUri);
    await vscode.window.showTextDocument(doc, { preview: false });
  };

  const handleOneDriveError = async (error: unknown): Promise<void> => {
    const message = error instanceof Error ? error.message : String(error);

    if (message.includes("AADSTS65002")) {
      const action = await vscode.window.showErrorMessage(
        "OneDrive Versions: Tenant policy blocked VS Code auth. Switch to device-code auth?",
        "Switch Auth Mode",
        "Open Setup Guide"
      );
      if (action === "Switch Auth Mode") {
        await client.setAuthMode("deviceCode");
        await vscode.commands.executeCommand("onedriveVersions.connectAccount");
      } else if (action === "Open Setup Guide") {
        await openSetupGuide();
      }
      return;
    }

    if (message.includes("auth.clientId")) {
      const action = await vscode.window.showErrorMessage(
        "OneDrive Versions: Device-code auth is not configured yet.",
        "Open Settings",
        "Open Setup Guide"
      );
      if (action === "Open Settings") {
        await vscode.commands.executeCommand("workbench.action.openSettings", "onedriveVersions.auth.clientId");
      } else if (action === "Open Setup Guide") {
        await openSetupGuide();
      }
      return;
    }

    void vscode.window.showErrorMessage(`OneDrive Versions: ${message}`);
  };

  const maybeShowFirstRunPrompt = async (): Promise<void> => {
    const shown = context.globalState.get<boolean>(onboardingKey, false);
    if (shown) {
      return;
    }

    const authMode = client.getAuthMode();
    if (authMode === "deviceCode" && !client.hasDeviceCodeClientId()) {
      const action = await vscode.window.showInformationMessage(
        "OneDrive Versions needs a Microsoft app client ID for device-code sign in.",
        "Open Settings",
        "Open Setup Guide"
      );
      if (action === "Open Settings") {
        await vscode.commands.executeCommand("workbench.action.openSettings", "onedriveVersions.auth.clientId");
      } else if (action === "Open Setup Guide") {
        await openSetupGuide();
      }
      await context.globalState.update(onboardingKey, true);
    }
  };

  const updateActiveContext = async (): Promise<void> => {
    const localPath = getActiveFilePath();
    const active = Boolean(localPath && client.findOneDriveRoot(localPath));
    await vscode.commands.executeCommand("setContext", "oneDriveVersions.active", active);

    if (!active || !localPath) {
      await vscode.commands.executeCommand("setContext", "oneDriveVersions.hasVersions", false);
      return;
    }

    const cached = client.getCachedContext(localPath);
    await vscode.commands.executeCommand("setContext", "oneDriveVersions.hasVersions", Boolean(cached?.versions.length));

    const autoLoad = vscode.workspace.getConfiguration("onedriveVersions").get<boolean>("autoLoadVersions", true);
    if (autoLoad && !cached) {
      try {
        await client.loadVersionsForFile(localPath, { interactive: false });
        await vscode.commands.executeCommand("setContext", "oneDriveVersions.hasVersions", true);
      } catch (error) {
        const msg = error instanceof Error ? error.message : String(error);
        if (msg === "AUTH_REQUIRED") {
          await vscode.commands.executeCommand("setContext", "oneDriveVersions.hasVersions", false);
          return;
        }
        if (!msg.includes("inside a detected OneDrive root")) {
          await handleOneDriveError(error);
        }
      }
    }
  };

  const openSelectedVersionPreview = async (localPath: string): Promise<void> => {
    const data = client.getCachedContext(localPath) ?? (await client.loadVersionsForFile(localPath));
    const version = data.versions[data.selectedIndex];
    if (!version) {
      throw new Error("No version selected.");
    }

    const fileName = path.basename(localPath);
    const uri = vscode.Uri.from({
      scheme: CONTENT_SCHEME,
      path: `/${fileName}`,
      query: `local=${encodeURIComponent(localPath)}`,
      fragment: `version=${encodeURIComponent(version.id)}`
    });

    const doc = await vscode.workspace.openTextDocument(uri);
    await vscode.window.showTextDocument(doc, { preview: true, preserveFocus: false });
  };

  const ensureVersions = async (localPath: string): Promise<VersionContext> => {
    const loaded = await client.loadVersionsForFile(localPath);
    await vscode.commands.executeCommand("setContext", "oneDriveVersions.hasVersions", loaded.versions.length > 0);
    return loaded;
  };

  const setSelectedIndex = async (localPath: string, nextIndex: number): Promise<void> => {
    const state = client.getCachedContext(localPath) ?? (await ensureVersions(localPath));
    if (!state.versions.length) {
      throw new Error("No versions available.");
    }
    const clamped = Math.max(0, Math.min(state.versions.length - 1, nextIndex));
    state.selectedIndex = clamped;
    await openSelectedVersionPreview(localPath);
  };

  context.subscriptions.push(
    vscode.commands.registerCommand("onedriveVersions.connectAccount", async () => {
      try {
        if (client.getAuthMode() === "deviceCode" && !client.hasDeviceCodeClientId()) {
          const action = await vscode.window.showWarningMessage(
            "Set onedriveVersions.auth.clientId before connecting with device code.",
            "Open Settings",
            "Open Setup Guide"
          );
          if (action === "Open Settings") {
            await vscode.commands.executeCommand("workbench.action.openSettings", "onedriveVersions.auth.clientId");
          } else if (action === "Open Setup Guide") {
            await openSetupGuide();
          }
          return;
        }

        await client.connectAccount();
        void vscode.window.showInformationMessage("OneDrive account connected.");
        await updateActiveContext();
      } catch (error) {
        await handleOneDriveError(error);
      }
    }),
    vscode.commands.registerCommand("onedriveVersions.openSetupGuide", async () => {
      await openSetupGuide();
    }),
    vscode.commands.registerCommand("onedriveVersions.pickVersion", async () => {
      const localPath = getActiveFilePath();
      if (!localPath) {
        void vscode.window.showInformationMessage("Open a file from a OneDrive folder first.");
        return;
      }

      try {
        const state = client.getCachedContext(localPath) ?? (await ensureVersions(localPath));
        const quickPickItems = state.versions.map((version, index) => {
          const modifiedBy = version.lastModifiedBy?.user?.displayName ?? "unknown";
          const dateString = new Date(version.lastModifiedDateTime).toLocaleString();
          const sizeString = typeof version.size === "number" ? `${Math.round(version.size / 1024)} KB` : "size n/a";
          return {
            label: `${state.selectedIndex === index ? "$(check) " : ""}${dateString}`,
            description: `${modifiedBy} | ${sizeString}`,
            detail: `Version ID: ${version.id}`,
            index
          };
        });

        const selected = await vscode.window.showQuickPick(quickPickItems, {
          title: "OneDrive Versions",
          placeHolder: "Choose a version to preview"
        });

        if (selected) {
          await setSelectedIndex(localPath, selected.index);
        }
      } catch (error) {
        await handleOneDriveError(error);
      }
    }),
    vscode.commands.registerCommand("onedriveVersions.previousVersion", async () => {
      const localPath = getActiveFilePath();
      if (!localPath) {
        return;
      }
      try {
        const state = client.getCachedContext(localPath) ?? (await ensureVersions(localPath));
        await setSelectedIndex(localPath, state.selectedIndex + 1);
      } catch (error) {
        await handleOneDriveError(error);
      }
    }),
    vscode.commands.registerCommand("onedriveVersions.nextVersion", async () => {
      const localPath = getActiveFilePath();
      if (!localPath) {
        return;
      }
      try {
        const state = client.getCachedContext(localPath) ?? (await ensureVersions(localPath));
        await setSelectedIndex(localPath, state.selectedIndex - 1);
      } catch (error) {
        await handleOneDriveError(error);
      }
    }),
    vscode.commands.registerCommand("onedriveVersions.saveAsVersion", async () => {
      const localPath = getActiveFilePath();
      if (!localPath) {
        return;
      }
      try {
        const state = client.getCachedContext(localPath) ?? (await ensureVersions(localPath));
        const selected = state.versions[state.selectedIndex];
        if (!selected) {
          throw new Error("No version selected.");
        }

        const targetUri = await vscode.window.showSaveDialog({
          title: "Save OneDrive Version As",
          defaultUri: vscode.Uri.file(localPath)
        });
        if (!targetUri) {
          return;
        }

        const bytes = await client.downloadVersionBytes(localPath, selected.id);
        await vscode.workspace.fs.writeFile(targetUri, bytes);
        void vscode.window.showInformationMessage(`Saved version ${selected.id} to ${targetUri.fsPath}`);
      } catch (error) {
        await handleOneDriveError(error);
      }
    }),
    vscode.commands.registerCommand("onedriveVersions.restoreVersion", async () => {
      const localPath = getActiveFilePath();
      if (!localPath) {
        return;
      }

      const activeDoc = vscode.window.activeTextEditor?.document;
      if (activeDoc?.isDirty && activeDoc.uri.scheme === "file" && samePath(activeDoc.uri.fsPath, localPath)) {
        void vscode.window.showWarningMessage("Save or discard local edits before restoring a OneDrive version.");
        return;
      }

      try {
        const state = client.getCachedContext(localPath) ?? (await ensureVersions(localPath));
        const selected = state.versions[state.selectedIndex];
        if (!selected) {
          throw new Error("No version selected.");
        }

        const confirm = await vscode.window.showWarningMessage(
          `Restore this file to the version from ${new Date(selected.lastModifiedDateTime).toLocaleString()}?`,
          { modal: true },
          "Restore"
        );

        if (confirm !== "Restore") {
          return;
        }

        const bytes = await client.downloadVersionBytes(localPath, selected.id);
        await vscode.workspace.fs.writeFile(vscode.Uri.file(localPath), bytes);

        const reopened = await vscode.workspace.openTextDocument(vscode.Uri.file(localPath));
        await vscode.window.showTextDocument(reopened, { preview: false });
        void vscode.window.showInformationMessage("OneDrive version restored locally. OneDrive sync will upload it as the current version.");
      } catch (error) {
        await handleOneDriveError(error);
      }
    })
  );

  context.subscriptions.push(
    vscode.window.onDidChangeActiveTextEditor(() => {
      void updateActiveContext();
    }),
    vscode.workspace.onDidCloseTextDocument((document) => {
      if (document.uri.scheme === "file") {
        client.clearCachedContext(document.uri.fsPath);
      }
    }),
    vscode.workspace.onDidChangeConfiguration((event) => {
      if (event.affectsConfiguration("onedriveVersions")) {
        void updateActiveContext();
      }
    })
  );

  void maybeShowFirstRunPrompt();
  void updateActiveContext();
}

export function deactivate(): void {
  // no-op
}

function getActiveFilePath(): string | undefined {
  const editor = vscode.window.activeTextEditor;
  if (!editor) {
    return undefined;
  }
  const uri = editor.document.uri;
  if (uri.scheme === "file") {
    return uri.fsPath;
  }
  if (uri.scheme === CONTENT_SCHEME) {
    const query = uri.query ?? "";
    const match = query.match(/(?:^|&)local=([^&]+)/);
    if (!match?.[1]) {
      return undefined;
    }
    return decodeURIComponent(match[1]);
  }
  return undefined;
}

function normalizeLocalRoot(input: string): string {
  return path.resolve(input).replace(/[\\/]+$/, "");
}

function normalizeRemoteRoot(input: string): string {
  const trimmed = input.trim().replace(/\\/g, "/");
  if (!trimmed || trimmed === "/") {
    return "/";
  }
  return `/${trimmed.replace(/^\/+/, "").replace(/\/+$/, "")}`;
}

function isPathWithin(candidate: string, root: string): boolean {
  const normalizedCandidate = normalizeLocalRoot(candidate);
  const normalizedRoot = normalizeLocalRoot(root);
  const relative = path.relative(normalizedRoot, normalizedCandidate);
  if (relative === "") {
    return true;
  }
  return !relative.startsWith("..") && !path.isAbsolute(relative);
}

function samePath(a: string, b: string): boolean {
  if (process.platform === "win32") {
    return normalizeLocalRoot(a).toLowerCase() === normalizeLocalRoot(b).toLowerCase();
  }
  return normalizeLocalRoot(a) === normalizeLocalRoot(b);
}

function isGraphNotFound(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error);
  return message.includes("Graph request failed (404)") || message.includes("itemNotFound");
}

function isGraphAccessDenied(error: unknown): boolean {
  const message = error instanceof Error ? error.message : String(error);
  return message.includes("Graph request failed (403)") || message.includes("accessDenied");
}

function buildRemotePathCandidates(remotePath: string): string[] {
  const normalized = remotePath.startsWith("/") ? remotePath : `/${remotePath}`;
  const segments = normalized.split("/").filter((s) => s.length > 0);
  if (!segments.length) {
    return ["/"];
  }

  const candidates: string[] = [];
  for (let start = 0; start < segments.length; start++) {
    candidates.push(`/${segments.slice(start).join("/")}`);
  }

  // De-duplicate while preserving order.
  return [...new Set(candidates)];
}

function normalizeShareBaseUrl(input: string): string {
  const trimmed = input.trim();
  if (!trimmed) {
    return trimmed;
  }

  // Some registry entries store malformed protocol like "https:/contoso...".
  const fixedProtocol = trimmed.replace(/^https:\/(?!\/)/i, "https://").replace(/^http:\/(?!\/)/i, "http://");
  return fixedProtocol;
}

function appendPathSegmentsToUrl(baseUrl: string, segments: string[]): string {
  const url = new URL(baseUrl);
  const basePath = url.pathname.replace(/\/+$/, "");
  const extraPath = segments.map((segment) => encodeURIComponent(segment)).join("/");
  url.pathname = extraPath ? `${basePath}/${extraPath}` : basePath || "/";
  return url.toString();
}

function toGraphShareId(webUrl: string): string {
  const base64 = Buffer.from(webUrl, "utf8").toString("base64");
  const base64Url = base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
  return `u!${base64Url}`;
}

function getRelativePathByUrlPrefix(targetUrl: string, baseUrl: string): string | undefined {
  try {
    const target = new URL(targetUrl);
    const base = new URL(baseUrl);

    if (target.origin.toLowerCase() !== base.origin.toLowerCase()) {
      return undefined;
    }

    const targetPath = target.pathname.replace(/\/+$/, "");
    const basePath = base.pathname.replace(/\/+$/, "");
    if (!targetPath.toLowerCase().startsWith(basePath.toLowerCase())) {
      return undefined;
    }

    const remaining = targetPath.slice(basePath.length).replace(/^\/+/, "");
    return decodeURIComponent(remaining);
  } catch {
    return undefined;
  }
}
