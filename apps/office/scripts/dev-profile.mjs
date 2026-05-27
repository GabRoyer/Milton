import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { homedir } from "node:os";
import { spawn, spawnSync } from "node:child_process";
import https from "node:https";
import net from "node:net";

const packageRoot = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const manifestTemplatePath = join(packageRoot, "manifests", "excel.local.xml");
const generatedRoot = join(packageRoot, ".generated");
const generatedManifestRoot = join(generatedRoot, "manifests");
const generatedProfilePath = join(generatedRoot, "dev-profile.json");
const registryPath = join(process.env.XDG_CONFIG_HOME ?? join(homedir(), ".config"), "milton", "office-dev-profiles.json");
const devNamespaceUuid = "3b2f8544-0f9e-4b9b-9864-14b2c6825457";
const defaultPortStart = 3100;
const defaultPortRange = 900;

try {
  await main();
} catch (error) {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
}

async function main() {
  const command = process.argv[2] ?? "generate";

  switch (command) {
    case "dev":
      await runDevServer();
      break;
    case "generate":
      await generateCommand();
      break;
    case "start":
      await officeAddinCommand("start");
      break;
    case "stop":
      await officeAddinCommand("stop");
      break;
    case "validate":
      await validateCommand();
      break;
    case "profiles":
      listProfiles();
      break;
    case "clean-profile":
      await cleanCurrentProfile();
      break;
    case "clean-stale-profiles":
      cleanStaleProfiles();
      break;
    default:
      console.error(`Unknown dev profile command: ${command}`);
      process.exit(1);
  }
}

async function runDevServer() {
  const profile = await prepareDevProfile();
  const portIsAvailable = await isPortAvailable(profile.port);

  if (!portIsAvailable) {
    if (await isMatchingDevServer(profile)) {
      console.log(`Milton Office dev server is already running for ${profile.displayName} at ${profile.origin}`);
      return;
    }

    console.error(
      `Milton dev port ${profile.port} is already in use for profile ${profile.shortProfile}. ` +
        "Stop the process using that port, clean the profile, or set MILTON_OFFICE_PORT.",
    );
    process.exit(1);
  }

  console.log(`Starting Milton Office dev server for ${profile.displayName} at ${profile.origin}`);
  await spawnPassthrough("vite", ["--host", "localhost", "--port", String(profile.port), "--strictPort"], profile);
}

async function generateCommand() {
  const profile = await prepareDevProfile();
  console.log(`Generated ${relative(packageRoot, profile.manifestPath)} for ${profile.displayName} at ${profile.origin}`);
}

async function officeAddinCommand(action) {
  const profile = await prepareDevProfile();
  const args = [action, relative(packageRoot, profile.manifestPath)];

  if (action === "start") {
    const portIsAvailable = await isPortAvailable(profile.port);

    if (!portIsAvailable && !(await isMatchingDevServer(profile))) {
      throw new Error(
        `Milton dev port ${profile.port} is already in use, but it does not look like profile ${profile.shortProfile}. ` +
          "Stop the process using that port, clean the profile, or set MILTON_OFFICE_PORT.",
      );
    }

    args.push("--dev-server-port", String(profile.port));
  }

  console.log(`${action === "start" ? "Starting" : "Stopping"} Milton Office add-in ${profile.displayName}`);
  await spawnPassthrough("office-addin-debugging", args, profile);
}

async function validateCommand() {
  const profile = await prepareDevProfile();
  await spawnPassthrough("office-addin-manifest", ["validate", relative(packageRoot, profile.manifestPath)], profile);
}

async function cleanCurrentProfile() {
  const profile = await resolveDevProfile({ allocatePort: false, persist: false });
  const registry = readRegistry();

  delete registry.profiles[profile.shortProfile];
  writeRegistry(registry);

  rmSync(profile.manifestPath, { force: true });

  if (existsSync(generatedProfilePath)) {
    const generatedProfile = readJson(generatedProfilePath);

    if (generatedProfile?.shortProfile === profile.shortProfile) {
      rmSync(generatedProfilePath, { force: true });
    }
  }

  console.log(`Removed Milton dev profile ${profile.shortProfile} for ${profile.worktreeRoot}`);
}

function listProfiles() {
  const registry = readRegistry();
  const entries = Object.entries(registry.profiles);

  if (entries.length === 0) {
    console.log("No Milton Office dev profiles are registered.");
    return;
  }

  for (const [shortProfile, profile] of entries) {
    const exists = existsSync(profile.worktreeRoot) ? "active" : "stale";
    console.log(`${shortProfile}  :${profile.port}  ${exists}  ${profile.worktreeRoot}`);
  }
}

function cleanStaleProfiles() {
  const registry = readRegistry();
  let removed = 0;

  for (const [shortProfile, profile] of Object.entries(registry.profiles)) {
    if (!existsSync(profile.worktreeRoot)) {
      delete registry.profiles[shortProfile];
      removed += 1;
    }
  }

  writeRegistry(registry);
  console.log(`Removed ${removed} stale Milton Office dev profile${removed === 1 ? "" : "s"}.`);
}

async function prepareDevProfile() {
  const profile = await resolveDevProfile({ allocatePort: true, persist: true });

  mkdirSync(generatedManifestRoot, { recursive: true });
  writeFileSync(profile.manifestPath, generateManifest(profile));
  writeFileSync(generatedProfilePath, `${JSON.stringify(profile, null, 2)}\n`);

  return profile;
}

async function resolveDevProfile(options) {
  const worktreeRoot = getWorktreeRoot();
  const manualProfile = process.env.MILTON_DEV_PROFILE?.trim();
  const profileSeed = manualProfile ? `${worktreeRoot}#${manualProfile}` : worktreeRoot;
  const profileKey = sha256(profileSeed);
  const shortProfile = profileKey.slice(0, 12);
  const manifestId = uuidV5(profileKey, devNamespaceUuid);
  const worktreeName = manualProfile || basenameForDisplay(worktreeRoot);
  const branch = gitOutput(["rev-parse", "--abbrev-ref", "HEAD"], worktreeRoot) || "unknown";
  const commit = gitOutput(["rev-parse", "--short", "HEAD"], worktreeRoot) || "unknown";
  const displayName = `Milton (${worktreeName}-${shortProfile.slice(0, 8)})`;
  const registry = readRegistry();
  const existing = registry.profiles[shortProfile];
  const envPort = parsePort(process.env.MILTON_OFFICE_PORT);
  const existingPort = existing?.worktreeRoot === worktreeRoot ? parsePort(existing.port) : undefined;
  const port = envPort ?? existingPort ?? (options.allocatePort ? await allocatePort(profileKey, registry, shortProfile) : defaultPort(profileKey));
  const origin = `https://localhost:${port}`;
  const manifestPath = join(generatedManifestRoot, `excel.${shortProfile}.xml`);
  const label = `Milton dev ${shortProfile.slice(0, 8)} | ${branch} | ${commit} | :${port}`;

  const profile = {
    profileKey,
    shortProfile,
    worktreeRoot,
    worktreeName,
    displayName,
    manifestId,
    port,
    origin,
    manifestPath,
    branch,
    commit,
    label,
    generatedAt: new Date().toISOString(),
  };

  if (options.persist) {
    registry.profiles[shortProfile] = {
      profileKey,
      worktreeRoot,
      displayName,
      manifestId,
      port,
      manifestPath,
      updatedAt: profile.generatedAt,
    };
    writeRegistry(registry);
  }

  return profile;
}

async function allocatePort(profileKey, registry, shortProfile) {
  let port = defaultPort(profileKey);

  for (let offset = 0; offset < defaultPortRange; offset += 1) {
    const candidate = defaultPortStart + ((port - defaultPortStart + offset) % defaultPortRange);
    const registered = Object.entries(registry.profiles).find(
      ([entryProfile, entry]) => entryProfile !== shortProfile && entry.port === candidate && existsSync(entry.worktreeRoot),
    );

    if (!registered && (await isPortAvailable(candidate))) {
      return candidate;
    }
  }

  throw new Error(`Unable to allocate a Milton Office dev port in ${defaultPortStart}-${defaultPortStart + defaultPortRange - 1}.`);
}

function defaultPort(profileKey) {
  return defaultPortStart + (Number.parseInt(profileKey.slice(0, 8), 16) % defaultPortRange);
}

function generateManifest(profile) {
  const displayName = escapeXmlAttribute(profile.displayName);
  const origin = profile.origin;

  return readFileSync(manifestTemplatePath, "utf8")
    .replace(/<Id>[^<]+<\/Id>/, `<Id>${profile.manifestId}</Id>`)
    .replace(/<DisplayName DefaultValue="[^"]*"\/>/, `<DisplayName DefaultValue="${displayName}"/>`)
    .replace(/https:\/\/localhost:\d+/g, origin);
}

async function spawnPassthrough(commandName, args, profile) {
  await new Promise((resolvePromise, rejectPromise) => {
    const child = spawn(commandName, args, {
      cwd: packageRoot,
      env: profileEnv(profile),
      stdio: "inherit",
    });

    child.on("error", rejectPromise);
    child.on("exit", (code, signal) => {
      if (signal) {
        rejectPromise(new Error(`${commandName} exited with signal ${signal}`));
        return;
      }

      if (code && code !== 0) {
        rejectPromise(new Error(`${commandName} exited with code ${code}`));
        return;
      }

      resolvePromise();
    });
  });
}

function profileEnv(profile) {
  return {
    ...process.env,
    MILTON_OFFICE_PORT: String(profile.port),
    MILTON_PUBLIC_DEV_PROFILE: profile.shortProfile,
    MILTON_PUBLIC_DEV_LABEL: profile.label,
    MILTON_PUBLIC_DEV_PORT: String(profile.port),
    MILTON_PUBLIC_DEV_BRANCH: profile.branch,
    MILTON_PUBLIC_DEV_COMMIT: profile.commit,
    MILTON_PUBLIC_DEV_WORKTREE: profile.worktreeName,
    npm_package_config_dev_server_port: String(profile.port),
  };
}

function isPortAvailable(port) {
  return new Promise((resolvePromise) => {
    const server = net.createServer();

    server.once("error", () => resolvePromise(false));
    server.once("listening", () => {
      server.close(() => resolvePromise(true));
    });
    server.listen({ host: "localhost", port });
  });
}

async function isMatchingDevServer(profile) {
  try {
    const source = await readHttpsText(`${profile.origin}/taskpanes/excel/main.tsx`);
    return source.includes("MILTON_PUBLIC_DEV_PROFILE") && source.includes(profile.shortProfile);
  } catch {
    return false;
  }
}

function readHttpsText(url) {
  return new Promise((resolvePromise, rejectPromise) => {
    const request = https.get(url, { rejectUnauthorized: false, timeout: 2000 }, (response) => {
      let body = "";

      response.setEncoding("utf8");
      response.on("data", (chunk) => {
        body += chunk;
      });
      response.on("end", () => {
        resolvePromise(body);
      });
    });

    request.on("timeout", () => {
      request.destroy(new Error(`Timed out reading ${url}`));
    });
    request.on("error", rejectPromise);
  });
}

function getWorktreeRoot() {
  const gitRoot = gitOutput(["rev-parse", "--show-toplevel"], packageRoot);

  if (!gitRoot) {
    throw new Error("Unable to resolve Git worktree root for Milton Office dev profile.");
  }

  return realpathSync(gitRoot);
}

function gitOutput(args, cwd) {
  const result = spawnSync("git", args, {
    cwd,
    encoding: "utf8",
    stdio: ["ignore", "pipe", "ignore"],
  });

  if (result.status !== 0) {
    return "";
  }

  return result.stdout.trim();
}

function readRegistry() {
  const registry = readJson(registryPath);

  return {
    version: 1,
    profiles: registry?.profiles && typeof registry.profiles === "object" ? registry.profiles : {},
  };
}

function writeRegistry(registry) {
  mkdirSync(dirname(registryPath), { recursive: true });
  writeFileSync(registryPath, `${JSON.stringify(registry, null, 2)}\n`);
}

function readJson(filePath) {
  if (!existsSync(filePath)) {
    return undefined;
  }

  try {
    return JSON.parse(readFileSync(filePath, "utf8"));
  } catch {
    return undefined;
  }
}

function parsePort(value) {
  if (value === undefined || value === null || value === "") {
    return undefined;
  }

  const port = Number(value);
  return Number.isInteger(port) && port > 0 && port < 65536 ? port : undefined;
}

function basenameForDisplay(filePath) {
  const normalized = filePath.replace(/\/+$/, "");
  const basename = normalized.split("/").pop();
  return basename || "worktree";
}

function sha256(value) {
  return createHash("sha256").update(value).digest("hex");
}

function uuidV5(name, namespaceUuid) {
  const namespace = Buffer.from(namespaceUuid.replace(/-/g, ""), "hex");
  const hash = createHash("sha1").update(namespace).update(name).digest();
  hash[6] = (hash[6] & 0x0f) | 0x50;
  hash[8] = (hash[8] & 0x3f) | 0x80;
  const hex = hash.subarray(0, 16).toString("hex");

  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}

function escapeXmlAttribute(value) {
  return value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
