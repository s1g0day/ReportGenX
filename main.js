const { app, BrowserWindow, ipcMain, shell, dialog, Menu } = require('electron')
const path = require('path')
const http = require('http')
const crypto = require('crypto')
const { spawn, execSync } = require('child_process')
const { loadSharedConfig, encodeBootstrapArgs } = require('./shared-config-utils')
const log = require('electron-log')

let pythonProcess = null
let mainWindow = null
let isQuitting = false
const APP_API_TOKEN = crypto.randomBytes(24).toString('hex')

const sharedConfig = loadSharedConfig(app, __dirname)
const SERVER_HOST = sharedConfig.server.host
const SERVER_PORT = sharedConfig.server.port
const ALLOWED_EXTERNAL_PROTOCOLS = new Set(sharedConfig.security.external_protocols)
const ALLOWED_EXTERNAL_HOSTS = new Set(sharedConfig.security.external_hosts)

function isExternalUrlAllowed(rawUrl) {
  try {
    const parsed = new URL(rawUrl)
    if (!ALLOWED_EXTERNAL_PROTOCOLS.has(parsed.protocol)) {
      return false
    }

    if (parsed.protocol === 'mailto:') {
      return true
    }

    return ALLOWED_EXTERNAL_HOSTS.has(parsed.hostname.toLowerCase())
  } catch (_e) {
    return false
  }
}

async function openExternalSafely(rawUrl, source) {
  if (!isExternalUrlAllowed(rawUrl)) {
    log.warn(`[Security] Blocked external URL from ${source}: ${rawUrl}`)
    return false
  }

  await shell.openExternal(rawUrl)
  return true
}

// 单实例锁：防止多次启动应用
const gotTheLock = app.requestSingleInstanceLock()

if (!gotTheLock) {
  // 如果获取锁失败，说明已有实例在运行，直接退出
  log.info('Another instance is already running. Exiting...')
  app.quit()
} else {
  // 当尝试启动第二个实例时，聚焦到第一个实例的窗口
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) {
        mainWindow.restore()
      }
      mainWindow.focus()
    }
  })
}

function startPythonBackend() {
  let backendExecutable = 'python'
  let args = ['-m', 'uvicorn', 'api:app', '--host', SERVER_HOST, '--port', String(SERVER_PORT)]
  let cwd = path.join(__dirname, 'backend')

  if (app.isPackaged) {
    // Production Mode - 目录模式：可执行文件在 dist/api/api (或 api.exe)
    const distPath = path.join(process.resourcesPath, 'backend', 'dist', 'api')
    cwd = path.join(process.resourcesPath, 'backend')

    if (process.platform === 'win32') {
      backendExecutable = path.join(distPath, 'api.exe')
    } else {
      backendExecutable = path.join(distPath, 'api')
    }
    args = [] // Executable handles main entry point
  } else {
    // Development Mode
    log.info(`Starting Python backend in dev mode: ${cwd}`)
  }

  pythonProcess = spawn(backendExecutable, args, {
    cwd,
    env: {
      ...process.env,
      PYTHONIOENCODING: 'utf-8',
      PYTHONUTF8: '1',
      APP_API_TOKEN
    }
  })

  pythonProcess.stdout.on('data', (data) => {
    log.info(`Python stdout: ${data.toString('utf-8')}`)
  })

  pythonProcess.stderr.on('data', (data) => {
    log.warn(`Python stderr: ${data.toString('utf-8')}`)
  })

  pythonProcess.on('error', (err) => {
    log.error(`Python process failed to start: ${err.message}`)
  })

  pythonProcess.on('close', (code) => {
    log.info(`Python process exited with code ${code}`)
    if (!isQuitting && code !== 0 && code !== null) {
      log.error(`Backend process exited unexpectedly with code ${code}`)
    }
  })
}

function checkBackendHealth() {
  return new Promise((resolve, reject) => {
    const req = http.request({
      hostname: SERVER_HOST,
      port: SERVER_PORT,
      path: '/api/health-auth',
      method: 'GET',
      headers: {
        'X-App-Token': APP_API_TOKEN
      }
    }, (res) => {
      if (res.statusCode && res.statusCode >= 200 && res.statusCode < 300) {
        resolve(true)
        return
      }

      if (res.statusCode === 403) {
        const mismatchError = new Error('Backend token mismatch')
        mismatchError.code = 'TOKEN_MISMATCH'
        reject(mismatchError)
        return
      }

      reject(new Error(`Health status ${res.statusCode}`))
    })

    req.on('error', reject)
    req.setTimeout(1500, () => {
      req.destroy(new Error('Health check timeout'))
    })
    req.end()
  })
}

async function waitForBackendReady(maxAttempts = 40, delayMs = 250) {
  for (let i = 0; i < maxAttempts; i += 1) {
    try {
      await checkBackendHealth()
      return
    } catch (e) {
      if (e && e.code === 'TOKEN_MISMATCH') {
        throw e
      }
      await new Promise((resolve) => setTimeout(resolve, delayMs))
    }
  }
  throw new Error(`Backend is not ready at http://${SERVER_HOST}:${SERVER_PORT}`)
}

function buildMenu() {
  const isMac = process.platform === 'darwin'

  const template = [
    // App 菜单 (macOS 独占)
    ...(isMac ? [{
      label: app.name,
      submenu: [
        { label: '关于 ReportGenX', role: 'about' },
        { type: 'separator' },
        { role: 'services' },
        { type: 'separator' },
        { role: 'hide' },
        { role: 'hideOthers' },
        { role: 'unhide' },
        { type: 'separator' },
        { role: 'quit' }
      ]
    }] : []),

    // 文件(&F)
    {
      label: '文件(&F)',
      submenu: [
        {
          label: '工具箱',
          accelerator: 'Ctrl+T',
          click: () => {
            if (mainWindow) {
              mainWindow.webContents.executeJavaScript('document.getElementById("btn-open-toolbox")?.click()')
            }
          }
        },
        {
          label: '生成报告',
          accelerator: 'Ctrl+Enter',
          click: () => {
            if (mainWindow) {
              mainWindow.webContents.executeJavaScript('document.getElementById("btn-dynamic-generate")?.click()')
            }
          }
        },
        { type: 'separator' },
        ...(isMac ? [] : [
          { label: '退出', accelerator: 'Ctrl+Q', role: 'quit' }
        ])
      ]
    },

    // 编辑(&E)
    {
      label: '编辑(&E)',
      submenu: [
        { label: '撤销', role: 'undo' },
        { label: '重做', role: 'redo' },
        { type: 'separator' },
        { label: '剪切', role: 'cut' },
        { label: '复制', role: 'copy' },
        { label: '粘贴', role: 'paste' },
        { label: '全选', role: 'selectAll' }
      ]
    },

    // 视图(&V)
    {
      label: '视图(&V)',
      submenu: [
        { label: '重新加载', role: 'reload' },
        { label: '强制重新加载', role: 'forceReload' },
        { label: '开发者工具', role: 'toggleDevTools' },
        { type: 'separator' },
        { label: '放大', role: 'zoomIn' },
        { label: '缩小', role: 'zoomOut' },
        { label: '重置缩放', role: 'resetZoom' }
      ]
    },

    // 帮助(&H)
    {
      label: '帮助(&H)',
      submenu: [
        ...(isMac ? [] : [
          { label: '关于 ReportGenX', role: 'about' },
          { type: 'separator' }
        ]),
        {
          label: '检查更新',
          click: () => {
            if (mainWindow) {
              mainWindow.webContents.executeJavaScript(`
                (async function () {
                  try {
                    var result = await AppAPI._request('/api/check-update')
                    if (result && result.has_update) {
                      var banner = document.getElementById('update-banner')
                      if (banner) {
                        banner.textContent = ''
                        banner.appendChild(document.createTextNode('新版本 '))
                        var strong = document.createElement('strong')
                        strong.textContent = result.latest_version
                        banner.appendChild(strong)
                        banner.appendChild(document.createTextNode(' 可用 (当前 ' + result.current_version + ') — '))
                        var link = document.createElement('a')
                        link.href = '#'
                        link.id = 'update-download-link'
                        link.textContent = '查看详情'
                        link.addEventListener('click', function (e) {
                          e.preventDefault()
                          if (window.electronAPI && window.electronAPI.openExternal) {
                            window.electronAPI.openExternal(result.download_url)
                          } else {
                            window.open(result.download_url, '_blank')
                          }
                        })
                        banner.appendChild(link)
                        banner.style.display = 'block'
                      }
                    } else {
                      if (window.AppUtils) AppUtils.showToast('当前已是最新版本', 'info')
                    }
                  } catch (e) {
                    if (window.AppUtils) AppUtils.showToast('检查更新失败', 'error')
                  }
                })()
              `)
            }
          }
        }
      ]
    }
  ]

  return Menu.buildFromTemplate(template)
}

function createWindow() {
  const bootstrapArgs = encodeBootstrapArgs(sharedConfig, APP_API_TOKEN)

  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      additionalArguments: bootstrapArgs,
      sandbox: true
    }
  })

  // 处理 window.open 跳转 (例如 target="_blank")
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    void openExternalSafely(url, 'window.open')
    return { action: 'deny' }
  })

  mainWindow.loadFile('src/index.html')
}

// IPC 处理
ipcMain.handle('open-external', async (event, url) => {
  const senderUrl = event.senderFrame && event.senderFrame.url
  if (typeof senderUrl !== 'string' || !senderUrl.startsWith('file://')) {
    throw new Error('Blocked by sender origin policy')
  }

  const opened = await openExternalSafely(url, 'ipc')
  if (!opened) {
    throw new Error('Blocked by external URL allowlist')
  }
  return { success: true }
})

app.whenReady().then(async () => {
  try {
    startPythonBackend()
    await waitForBackendReady()
    Menu.setApplicationMenu(buildMenu())
    createWindow()

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createWindow()
      }
    })
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err)
    if (err && err.code === 'TOKEN_MISMATCH') {
      dialog.showErrorBox('启动失败', '检测到端口上已有其他后端实例（令牌不匹配）。请关闭旧的 ReportGenX 后端进程后重试。')
    } else {
      dialog.showErrorBox('启动失败', `后端服务未就绪：${message}`)
    }
    app.quit()
  }
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('will-quit', () => {
  isQuitting = true
  if (pythonProcess) {
    if (process.platform === 'win32') {
      try {
        execSync(`taskkill /pid ${pythonProcess.pid} /T /F`, { stdio: 'ignore', timeout: 3000 })
        log.info(`Backend process tree killed (PID ${pythonProcess.pid})`)
      } catch (_e) {
        log.warn(`taskkill failed for PID ${pythonProcess.pid}, process may have already exited`)
      }
    } else {
      try {
        process.kill(-pythonProcess.pid, 'SIGKILL')
      } catch (_e) {
        pythonProcess.kill('SIGKILL')
      }
    }
    pythonProcess = null
  }
})
