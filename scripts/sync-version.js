/**
 * 版本同步脚本
 * - package.json -> backend/config.yaml (version: Vx.y.z)
 * - package.json -> backend/shared-config.json (app.version: x.y.z)
 * - package.json -> README.md (当前归档版本：`x.y.z`)
 */

const fs = require('fs');
const path = require('path');

const ROOT_DIR = path.join(__dirname, '..');
const PACKAGE_JSON = path.join(ROOT_DIR, 'package.json');
const CONFIG_YAML = path.join(ROOT_DIR, 'backend', 'config.yaml');
const SHARED_CONFIG_JSON = path.join(ROOT_DIR, 'backend', 'shared-config.json');
const README_MD = path.join(ROOT_DIR, 'README.md');

const DEFAULT_SHARED_CONFIG = {
  server: {
    host: '127.0.0.1',
    port: 8000,
  },
  app: {
    version: '0.0.0',
  },
  security: {
    external_protocols: ['https:'],
    external_hosts: ['github.com', 'www.github.com'],
  },
  paths: {
    open_folder_allowlist: ['output/report', 'output/temp', 'output'],
  },
  plugin_runtime: {
    mode: 'hybrid',
    use_legacy_core_alias: false,
    force_legacy_templates: [],
    subprocess_strategy: 'hybrid',
    subprocess_timeout_seconds: 120,
    isolated_enabled_templates: [],
    isolated_disabled_templates: [],
    isolated_rollout_percent: 0,
    isolated_template_rollout: {},
    isolated_fallback_mode: 'hybrid',
    metrics_emit_every_n: 50,
  },
};

function readJson(filePath, fallback = {}) {
  if (!fs.existsSync(filePath)) {
    return fallback;
  }

  try {
    return JSON.parse(fs.readFileSync(filePath, 'utf-8'));
  } catch (error) {
    console.warn(`[sync-version] Invalid JSON at ${filePath}, using fallback.`);
    return fallback;
  }
}

function syncConfigYaml(version) {
  let configContent = fs.readFileSync(CONFIG_YAML, 'utf-8');

  const versionRegex = /^version:\s*V?[\d.]+/m;
  const newVersion = `version: V${version}`;

  const currentMatch = configContent.match(versionRegex);
  if (!currentMatch) {
    throw new Error('version field not found in backend/config.yaml');
  }

  // Skip write if version already matches (avoids unnecessary file modification on every start)
  if (currentMatch[0] === newVersion) {
    return;
  }

  configContent = configContent.replace(versionRegex, newVersion);
  fs.writeFileSync(CONFIG_YAML, configContent, 'utf-8');
}

function syncSharedConfig(version) {
  const current = readJson(SHARED_CONFIG_JSON, DEFAULT_SHARED_CONFIG);

  // Skip write if version already matches
  if (current.app && current.app.version === version) {
    return;
  }

  const next = {
    ...DEFAULT_SHARED_CONFIG,
    ...current,
    server: {
      ...DEFAULT_SHARED_CONFIG.server,
      ...(current.server || {}),
    },
    app: {
      ...DEFAULT_SHARED_CONFIG.app,
      ...(current.app || {}),
      version,
    },
    security: {
      ...DEFAULT_SHARED_CONFIG.security,
      ...(current.security || {}),
    },
    paths: {
      ...DEFAULT_SHARED_CONFIG.paths,
      ...(current.paths || {}),
    },
    plugin_runtime: {
      ...DEFAULT_SHARED_CONFIG.plugin_runtime,
      ...(current.plugin_runtime || {}),
    },
  };

  fs.writeFileSync(SHARED_CONFIG_JSON, `${JSON.stringify(next, null, 2)}\n`, 'utf-8');
}

function syncReadme(version) {
  let readmeContent = fs.readFileSync(README_MD, 'utf-8');

  const versionRegex = /当前归档版本：`[\d.]+`/;
  const newVersion = `当前归档版本：\`${version}\``;

  const currentMatch = readmeContent.match(versionRegex);
  if (!currentMatch) {
    throw new Error('version line not found in README.md');
  }

  // Skip write if version already matches
  if (currentMatch[0] === newVersion) {
    return;
  }

  readmeContent = readmeContent.replace(versionRegex, newVersion);
  fs.writeFileSync(README_MD, readmeContent, 'utf-8');
}

function syncVersion() {
  const packageJson = readJson(PACKAGE_JSON);
  if (!packageJson.version) {
    throw new Error('version not found in package.json');
  }

  const version = packageJson.version;
  syncConfigYaml(version);
  syncSharedConfig(version);
  syncReadme(version);

  console.log(`✓ Version synced: ${version}`);
}

syncVersion();
