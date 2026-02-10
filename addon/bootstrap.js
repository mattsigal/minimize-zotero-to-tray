// Zotero-in-Tray with TCP Helper Process
// This version uses an external AutoHotkey executable for the tray icon
// to avoid js-ctypes crashes on Windows 11. Communication is via TCP sockets.

var ZoteroInTray = {
    // Plugin metadata
    id: null,
    version: null,
    rootURI: null,

    // Logging function, initialized immediately.
    log: (msg) => {
        if (typeof Zotero !== 'undefined' && Zotero.debug) {
            Zotero.debug(`ZoteroInTray: ${msg}`);
        } else {
            console.log(`ZoteroInTray: ${msg}`);
        }
    },

    // For TCP Server
    serverSocket: null,

    // Window Management
    mainWindow: null,
    mainWindowHandle: null,
    lockedWindows: new Map(),
    isWindowHidden: false,
    windowWasMaximized: false, // Reverted to simple boolean logic
    isActuallyQuitting: false,
    initialHidePerformed: false,
    hidePollingInterval: null,

    // Helper Process
    helperProcess: null,
    helperExeName: 'tray_helper.exe',
    helperPath: null,
    vbsPath: null,
    isShuttingDown: false,
    relaunchDelay: 2000, // ms
    prefObserver: null,

    // Windows API
    user32: null,
    kernel32: null,
    ctypes: null,

    // WinAPI Constants
    constants: {
        SW_HIDE: 0,
        SW_RESTORE: 9,
        SW_MAXIMIZE: 3,
        SW_MINIMIZE: 6,
    },

    // Mozilla Components
    Cc: null,
    Ci: null,
    prefPane: null,

    init: function ({ id, version, rootURI }) {
        this.id = id;
        this.version = version;
        this.rootURI = rootURI;

        // Reset state for re-enabling plugin without Zotero restart
        this.isShuttingDown = false;
        this.isShuttingDown = false;
        this.initialHidePerformed = false;
        this.cleanupPerformed = false; // Reset cleanup flag

        try {
            // Define core components
            this.Cc = Components.classes;
            this.Ci = Components.interfaces;

            // Import ctypes with fallback for Zotero 7 / Firefox 115+
            try {
                // Try standard JSM first (Zotero 6 / Early 7)
                const { ctypes } = ChromeUtils.import("resource://gre/modules/ctypes.jsm");
                this.ctypes = ctypes;
            } catch (e) {
                this.log("⚠️ Standard ctypes.jsm import failed. Trying ES Module...");
                try {
                     // Try ES Module (Modern Firefox/Zotero 7+)
                     const { ctypes } = ChromeUtils.importESModule("resource://gre/modules/ctypes.sys.mjs");
                     this.ctypes = ctypes;
                } catch (e2) {
                     this.log(`FATAL: Failed to import ctypes via JSM or ESM: ${e2}`);
                     return;
                }
            }

        } catch (e) {
            this.log(`FATAL: Failed to import critical JSMs: ${e}`);
            return; // Abort initialization
        }

        this.log('🚀 Initializing Zotero-in-Tray (TCP Architecture)...');

        this.initWinAPI();
        this.startServer();
        this.registerPrefObserver();
        this.registerPreferences();
        this.launchHelper();
        this.setupDualInterceptForExistingWindows();

        this.log("✓ Initialization complete.");
    },

    initWinAPI: function () {
        if (!this.ctypes) {
            this.log("✗ ctypes not available");
            return;
        }
        try {
            this.log("Initializing Windows API libraries...");
            this.user32 = this.ctypes.open("user32.dll");
            this.kernel32 = this.ctypes.open("kernel32.dll");
            this.log("✓ Windows API libraries loaded.");
            this.declareWinAPIFunctions();
        } catch (e) {
            this.log("✗ Error initializing Windows API: " + e);
        }
    },

    startServer: function () {
        try {
            const port = Zotero.Prefs.get('extensions.zotero-in-tray.network.port', true);

            // FORCE ENABLE AUTO-HIDE (Safely)
            try {
                Zotero.Prefs.set('extensions.zotero-in-tray.startup.autohide', true, true);
                this.log('✓ Enforcing "Auto-hide on startup" preference (Safe Mode).');
            } catch (e) {
                this.log('⚠️ Could not force auto-hide pref: ' + e);
            }

            this.log(`Attempting to start server on port: ${port} (Type: ${typeof port})`);

            if (!port || isNaN(port)) {
                this.log(`✗ Invalid port number: '${port}'. Aborting server start.`);
                return;
            }

            this.serverSocket = this.Cc["@mozilla.org/network/server-socket;1"]
                .createInstance(this.Ci.nsIServerSocket);

            const listener = {
                onSocketAccepted: (socket, transport) => {
                    this.log("TCP Server: Connection accepted.");
                    this.handleConnection(socket, transport.openInputStream(0, 0, 0), transport.openOutputStream(0, 0, 0));
                }
            };
            this.serverSocket.init(Number(port), true, -1);
            this.serverSocket.asyncListen(listener);
            this.log(`✓ Server listening on port ${port}`);
        } catch (e) {
            this.log(`✗ Error starting server: ${e}`);
            if (typeof Zotero !== 'undefined') Zotero.logError(e);
        }
    },

    declareWinAPIFunctions: function () {
        try {
            this.user32.FindWindowW = this.user32.declare("FindWindowW", this.ctypes.winapi_abi, this.ctypes.voidptr_t, this.ctypes.char16_t.ptr, this.ctypes.char16_t.ptr);
            this.user32.ShowWindow = this.user32.declare("ShowWindow", this.ctypes.winapi_abi, this.ctypes.bool, this.ctypes.voidptr_t, this.ctypes.int);
            this.user32.SetForegroundWindow = this.user32.declare("SetForegroundWindow", this.ctypes.winapi_abi, this.ctypes.bool, this.ctypes.voidptr_t);
            this.user32.IsWindowVisible = this.user32.declare("IsWindowVisible", this.ctypes.winapi_abi, this.ctypes.bool, this.ctypes.voidptr_t);
            this.user32.IsZoomed = this.user32.declare("IsZoomed", this.ctypes.winapi_abi, this.ctypes.bool, this.ctypes.voidptr_t);
            this.user32.GetForegroundWindow = this.user32.declare("GetForegroundWindow", this.ctypes.winapi_abi, this.ctypes.voidptr_t);
            this.user32.IsIconic = this.user32.declare("IsIconic", this.ctypes.winapi_abi, this.ctypes.bool, this.ctypes.voidptr_t);

            // Functions for robust focus handling
            this.kernel32.GetCurrentThreadId = this.kernel32.declare("GetCurrentThreadId", this.ctypes.winapi_abi, this.ctypes.uint32_t);
            this.user32.GetWindowThreadProcessId = this.user32.declare("GetWindowThreadProcessId", this.ctypes.winapi_abi, this.ctypes.uint32_t, this.ctypes.voidptr_t, this.ctypes.voidptr_t);
            this.user32.AttachThreadInput = this.user32.declare("AttachThreadInput", this.ctypes.winapi_abi, this.ctypes.bool, this.ctypes.uint32_t, this.ctypes.uint32_t, this.ctypes.bool);

            // PID-based window finding functions
            this.kernel32.GetCurrentProcessId = this.kernel32.declare("GetCurrentProcessId", this.ctypes.winapi_abi, this.ctypes.uint32_t);

            this.log("✓ Windows API functions declared.");
        } catch (e) {
            this.log("✗ Error declaring Windows API functions: " + e);
            throw e;
        }
    },

    getHotkeyArgs: function () {
        const args = [];
        try {
            const useCtrl = Zotero.Prefs.get('extensions.zotero-in-tray.hotkey.ctrl', true);
            const useAlt = Zotero.Prefs.get('extensions.zotero-in-tray.hotkey.alt', true);
            const useShift = Zotero.Prefs.get('extensions.zotero-in-tray.hotkey.shift', true);
            const key = Zotero.Prefs.get('extensions.zotero-in-tray.hotkey.key', true);
            const port = Zotero.Prefs.get('extensions.zotero-in-tray.network.port', true);
            this.log(`Read port for helper args: ${port}`);

            if (useCtrl) args.push('--ctrl');
            if (useAlt) args.push('--alt');
            if (useShift) args.push('--shift');

            if (key && /^[a-zA-Z0-9]$/.test(key)) {
                args.push(`--key=${key.toUpperCase()}`);
            } else if (key) {
                this.log(`✗ Invalid hotkey character specified: "${key}". Ignoring.`);
            }

            if (port && !isNaN(port)) {
                args.push(`--port=${port}`);
            } else {
                this.log(`✗ Invalid or missing port for helper. Using helper's default.`);
            }
        } catch (e) {
            this.log(`✗ Error reading preferences for helper: ${e}`);
        }
        return args;
    },

    launchHelper: function () {
        this.log("🚀 launchHelper called!"); // PROOF OF LIFE
        if (this.isShuttingDown) {
            this.log("Shutdown in progress, aborting helper launch.");
            return;
        }
        this.log("🚀 Launching helper process...");
        try {
            let bytes;
            if (this.rootURI.startsWith("jar:")) {
                const jarPath = this.rootURI.substring(4, this.rootURI.indexOf('!'));
                const fileHandler = this.Cc["@mozilla.org/network/protocol;1?name=file"].getService(this.Ci.nsIFileProtocolHandler);
                const xpiFile = fileHandler.getFileFromURLSpec(jarPath);

                const zr = this.Cc["@mozilla.org/libjar/zip-reader;1"].createInstance(this.Ci.nsIZipReader);
                zr.open(xpiFile);

                const entryPath = "bin/" + this.helperExeName;
                if (!zr.hasEntry(entryPath)) {
                    zr.close();
                    throw new Error(`Helper executable not found in XPI at path: ${entryPath}`);
                }
                const inputStream = zr.getInputStream(entryPath);
                const binaryInputStream = this.Cc["@mozilla.org/binaryinputstream;1"].createInstance(this.Ci.nsIBinaryInputStream);
                binaryInputStream.setInputStream(inputStream);
                bytes = binaryInputStream.readBytes(binaryInputStream.available());
                zr.close();
            } else {
                // Handle unpacked directory (Zotero 8+ typical behavior)
                const fileHandler = this.Cc["@mozilla.org/network/protocol;1?name=file"].getService(this.Ci.nsIFileProtocolHandler);
                const extDir = fileHandler.getFileFromURLSpec(this.rootURI);
                const sourceExe = extDir.clone();
                sourceExe.append("bin");
                sourceExe.append(this.helperExeName);

                this.log(`Loading helper from file path: ${sourceExe.path}`);
                const fis = this.Cc["@mozilla.org/network/file-input-stream;1"].createInstance(this.Ci.nsIFileInputStream);
                fis.init(sourceExe, 0x01, 0o444, 0);
                const bis = this.Cc["@mozilla.org/binaryinputstream;1"].createInstance(this.Ci.nsIBinaryInputStream);
                bis.setInputStream(fis);
                bytes = bis.readBytes(bis.available());
                fis.close();
            }

            const dirService = this.Cc['@mozilla.org/file/directory_service;1'].getService(this.Ci.nsIDirectoryService);
            const tmpDir = dirService.get("TmpD", this.Ci.nsIFile);

            const helperFile = tmpDir.clone();
            helperFile.append(this.helperExeName);
            this.helperPath = helperFile.path;

            const ostream = this.Cc["@mozilla.org/network/file-output-stream;1"].createInstance(this.Ci.nsIFileOutputStream);
            ostream.init(helperFile, 0x02 | 0x08 | 0x20, 0o755, 0);
            ostream.write(bytes, bytes.length);
            ostream.close();
            this.log(`✓ Helper extracted to: ${this.helperPath}`);

            // NEW: Also extract the icon file!
            try {
                const iconName = "zotero_128.ico";
                const iconEntry = "bin/" + iconName;

                // Re-open zip or standard copy depending on mode
                let iconBytes;
                if (this.rootURI.startsWith("jar:")) {
                    const jarPath2 = this.rootURI.substring(4, this.rootURI.indexOf('!'));
                    const fileHandler2 = this.Cc["@mozilla.org/network/protocol;1?name=file"].getService(this.Ci.nsIFileProtocolHandler);
                    const xpiFile2 = fileHandler2.getFileFromURLSpec(jarPath2);
                    const zr2 = this.Cc["@mozilla.org/libjar/zip-reader;1"].createInstance(this.Ci.nsIZipReader);
                    zr2.open(xpiFile2);
                    if (zr2.hasEntry(iconEntry)) {
                        const is2 = zr2.getInputStream(iconEntry);
                        const bis2 = this.Cc["@mozilla.org/binaryinputstream;1"].createInstance(this.Ci.nsIBinaryInputStream);
                        bis2.setInputStream(is2);
                        iconBytes = bis2.readBytes(bis2.available());
                    }
                    zr2.close();
                } else {
                    const fh2 = this.Cc["@mozilla.org/network/protocol;1?name=file"].getService(this.Ci.nsIFileProtocolHandler);
                    const ed2 = fh2.getFileFromURLSpec(this.rootURI);
                    const srcIcon = ed2.clone();
                    srcIcon.append("bin");
                    srcIcon.append(iconName);
                    if (srcIcon.exists()) {
                        const fis2 = this.Cc["@mozilla.org/network/file-input-stream;1"].createInstance(this.Ci.nsIFileInputStream);
                        fis2.init(srcIcon, 0x01, 0o444, 0);
                        const bis2 = this.Cc["@mozilla.org/binaryinputstream;1"].createInstance(this.Ci.nsIBinaryInputStream);
                        bis2.setInputStream(fis2);
                        iconBytes = bis2.readBytes(bis2.available());
                        fis2.close();
                    }
                }

                if (iconBytes) {
                    const iconDest = tmpDir.clone();
                    iconDest.append(iconName);
                    const os2 = this.Cc["@mozilla.org/network/file-output-stream;1"].createInstance(this.Ci.nsIFileOutputStream);
                    os2.init(iconDest, 0x02 | 0x08 | 0x20, 0o644, 0);
                    os2.write(iconBytes, iconBytes.length);
                    os2.close();
                    this.log(`✓ Icon extracted to: ${iconDest.path}`);
                } else {
                    this.log("⚠️ Could not find icon file to extract.");
                }
            } catch (e) {
                this.log("⚠️ Icon extraction failed (non-fatal): " + e);
            }

            // [NEW] Create silent_kill.vbs for Flashbang-free shutdown
            try {
                const vbsName = "silent_kill.vbs";
                const vbsFile = tmpDir.clone();
                vbsFile.append(vbsName);
                this.vbsPath = vbsFile.path;

                // VBS Logic: Create WScript.Shell and Run taskkill hidden (0)
                const vbsContent = 'CreateObject("WScript.Shell").Run "taskkill /F /IM tray_helper.exe", 0';

                const os3 = this.Cc["@mozilla.org/network/file-output-stream;1"].createInstance(this.Ci.nsIFileOutputStream);
                os3.init(vbsFile, 0x02 | 0x08 | 0x20, 0o644, 0);
                os3.write(vbsContent, vbsContent.length);
                os3.close();
                this.log(`✓ Silent killer script created: ${this.vbsPath}`);
            } catch (e) {
                this.log("⚠️ VBS creation failed: " + e);
            }

            const process = this.Cc["@mozilla.org/process/util;1"].createInstance(this.Ci.nsIProcess);
            process.init(helperFile);

            const args = this.getHotkeyArgs();
            this.log(`🚀 Running helper with args: ${args.join(' ')}`);
            process.runAsync(args, args.length, (subject, topic, data) => {
                if (topic === "process-finished" || topic === "process-failed") {
                    this.log(`Helper process terminated (topic: ${topic}). Exit code: ${data}`);
                    this.helperProcess = null;
                    if (!this.isShuttingDown) {
                        this.log(`🤔 Helper process terminated unexpectedly. Restarting in ${this.relaunchDelay / 1000}s...`);
                        setTimeout(() => {
                            this.log("Attempting to relaunch helper process...");
                            this.launchHelper();
                        }, this.relaunchDelay);
                    }
                }
            });
            this.helperProcess = process;

            try {
                // FORCE POLLING - Bypass Pref check to guarantee run
                // const shouldAutoHide = Zotero.Prefs.get('extensions.zotero-in-tray.startup.autohide', true);

                if (!this.initialHidePerformed) {
                    this.log('🚀 Auto-hide logic STARTED. Polling for window...');
                    if (this.hidePollingInterval) clearInterval(this.hidePollingInterval);

                    this.hidePollingInterval = setInterval(() => {
                        this.tryHideWindowOnStartup();
                    }, 500); // Increased to 500ms to be nicer to CPU

                    setTimeout(() => {
                        if (this.hidePollingInterval) {
                            clearInterval(this.hidePollingInterval);
                            this.hidePollingInterval = null;
                            this.log('✗ Polling for window handle timed out after 15s.');
                        }
                    }, 15000);
                }
            } catch (e) {
                this.log(`✗ Error starting auto-hide poller: ${e}`);
            }

        } catch (e) {
            this.log("✗ Error launching helper process: " + e);
            if (typeof Zotero !== 'undefined') Zotero.logError(e);
        }
    },

    registerPreferences: function () {
        this.log("Registering preferences pane...");
        this.prefPane = Zotero.PreferencePanes.register({
            pluginID: this.id,
            paneID: 'zotero-in-tray-prefs',
            label: 'Minimize to Tray',
            src: this.rootURI + 'preferences.xhtml',
        });
        this.log("✓ Preferences pane registered.");
    },

    registerPrefObserver: function () {
        this.log("Registering preference observer...");

        this.prefObserver = (branch, name) => {
            if (name.startsWith('extensions.zotero-in-tray.')) {
                this.log(`Preference changed: ${name}. Restarting helper process.`);
                if (this.helperProcess) {
                    this.helperProcess.kill();
                } else {
                    this.log("Helper process was not running, launching it now.");
                    this.launchHelper();
                }
            }
        };

        Zotero.Prefs.registerObserver('extensions.zotero-in-tray.', this.prefObserver);
        this.log("✓ Preference observer registered.");
    },

    handleConnection: function (socket, inputStream, outputStream) {
        this.log('✓ Client connection accepted. Setting up data pump...');
        try {
            const pump = this.Cc['@mozilla.org/network/input-stream-pump;1'].createInstance(this.Ci.nsIInputStreamPump);
            pump.init(inputStream, -1, -1, true);

            const listener = {
                onStartRequest: (request) => { this.log('Pump: onStartRequest'); },
                onStopRequest: (request, statusCode) => { this.log(`Pump: onStopRequest. Status: ${statusCode}`); },
                onDataAvailable: (request, stream, offset, count) => {
                    try {
                        const scriptableStream = this.Cc['@mozilla.org/scriptableinputstream;1'].createInstance(this.Ci.nsIScriptableInputStream);
                        scriptableStream.init(stream);
                        const data = scriptableStream.read(count);
                        this.log(`📥 Received command: ${data}`);

                        if (data.trim() === 'CLICKED') {
                            const threadManager = this.Cc["@mozilla.org/thread-manager;1"].getService(this.Ci.nsIThreadManager);
                            threadManager.mainThread.dispatch(() => {
                                this.handleTrayClick();
                            }, this.Ci.nsIThread.DISPATCH_NORMAL);
                        }
                    } catch (e) {
                        this.log(`✗ Error in onDataAvailable: ${e}`);
                    }
                }
            };

            pump.asyncRead(listener, null);
            this.log('✓ Pump configured and asyncRead called.');

        } catch (e) {
            this.log(`✗ Error setting up pump: ${e}`);
        }
    },

    getMainWindowHandle: function () {
        if (this.mainWindowHandle && !this.mainWindowHandle.isNull()) return true;

        // PRIORITY 0: God Mode (Internal Mozilla API)
        try {
            if (this.mainWindow && this.mainWindow.docShell) {
                const baseWindow = this.mainWindow.docShell.treeOwner.QueryInterface(this.Ci.nsIBaseWindow);
                const nativeHandleString = baseWindow.nativeHandle;

                if (nativeHandleString) {
                    let handleInt = BigInt(nativeHandleString);
                    this.mainWindowHandle = this.ctypes.voidptr_t(handleInt.toString());
                    this.log("✅ SUCCESS: Acquired Native Handle via nsIBaseWindow: " + this.mainWindowHandle.toString());
                    return true;
                }
            }
        } catch (e) {
            this.log(`⚠️ nsIBaseWindow method failed: ${e}`);
        }

        this.log("🔍 Attempting PID-based window finding (Fallback)...");
        try {
            const currentPID = this.kernel32.GetCurrentProcessId();

            // PRIORITY 1: Check Foreground Window (Most likely Zotero on startup)
            const fgHandle = this.user32.GetForegroundWindow();
            if (fgHandle && !fgHandle.isNull()) {
                const processIdPtr = this.ctypes.uint32_t();
                this.user32.GetWindowThreadProcessId(fgHandle, processIdPtr.address());
                if (processIdPtr.value === currentPID) {
                    this.mainWindowHandle = fgHandle;
                    this.log("✅ SUCCESS: Found window handle via Foreground check!");
                    return true;
                }
            }

            // PRIORITY 2: Check standard class name
            const windowClasses = ["MozillaWindowClass"];
            for (const className of windowClasses) {
                const handle = this.user32.FindWindowW(this.ctypes.char16_t.array()(className), null);
                if (handle && !handle.isNull()) {
                    const processIdPtr = this.ctypes.uint32_t();
                    this.user32.GetWindowThreadProcessId(handle, processIdPtr.address());
                    if (processIdPtr.value === currentPID) {
                        this.mainWindowHandle = handle;
                        this.log("✅ SUCCESS: Found window handle via Class Name.");
                        return true;
                    }
                }
            }
        } catch (e) {
            this.log(`❌ Error during window finding: ${e}`);
        }
        return false;
    },



    setupDualInterceptForExistingWindows: function () {
        this.log("🔥 Setting up DUAL INTERCEPT for existing windows...");
        let mainWindows = Zotero.getMainWindows();
        for (let window of mainWindows) {
            this.lockWindow(window);
        }
        if (mainWindows.length > 0) {
            this.mainWindow = mainWindows[0];
            // Only try to get handle if we have a valid Zotero window
            if (this.mainWindow) {
                this.getMainWindowHandle();
            }
            this.log(`✓ DUAL INTERCEPT set up for ${mainWindows.length} Zotero windows`);
        } else {
            this.log("No existing Zotero windows found, will wait for onMainWindowLoad");
        }
    },

    lockWindow: function (window) {
        if (!window || this.lockedWindows.has(window)) return;
        this.log("🔒 Locking window: " + window.location.href);

        try {
            let self = this;

            // Minimize Handler (Keep this!)
            let minimizeHandler = function (event) {
                if (window.windowState === 2) { // 2 = STATE_MINIMIZED
                    self.log("🔥🔥 MINIMIZE EVENT detected! Hiding window to tray (delayed).");
                    // Delay to let Windows finish the minimize animation/state change
                    setTimeout(() => {
                        self.hideMainWindow();
                    }, 200);
                }
            };

            // ONLY intercept minimize (sizemodechange). Let 'close' happen naturally.
            window.addEventListener("sizemodechange", minimizeHandler, false);

            this.lockedWindows.set(window, { minimizeHandler });
            this.log("✓ Window locked with Single intercept (Minimize Only)");
        } catch (e) {
            this.log("✗ Failed to lock window: " + e);
        }
    },

    unlockWindow: function (window) {
        if (!window || !this.lockedWindows.has(window)) return;
        try {
            this.log("🔓 Unlocking window: " + window.location.href);
            let lockInfo = this.lockedWindows.get(window);

            if (lockInfo.minimizeHandler) {
                window.removeEventListener("sizemodechange", lockInfo.minimizeHandler, false);
            }
            // No close handler to remove anymore

            this.lockedWindows.delete(window);
            this.log("✓ Window unlocked");
        } catch (e) {
            this.log("✗ Failed to unlock window: " + e);
        }
    },

    onWindowClosing: function () {
        this.hideMainWindow();
    },

    handleTrayClick: function () {
        this.log('🖱️ Tray icon/hotkey handled.');
        try {
            if (!this.getMainWindowHandle()) {
                this.log("✗ Could not get main window handle for tray click.");
                return;
            }

            const isVisible = this.user32.IsWindowVisible(this.mainWindowHandle);
            const isIconic = this.user32.IsIconic(this.mainWindowHandle);
            const isForeground = this.user32.GetForegroundWindow().toString() === this.mainWindowHandle.toString();

            this.log(`Window state: isVisible=${isVisible}, isIconic=${isIconic}, isForeground=${isForeground}`);

            if (isIconic) {
                // Case 1: Window is minimized to the taskbar. Restore it intelligently.
                this.log("🔄 Window is minimized, restoring...");
                this.showMainWindow({ forceRestore: true });
            } else if (!isVisible) {
                // Case 2: Window was hidden by us. Restore using the saved state.
                this.log("🔄 Window is hidden by plugin, showing...");
                this.showMainWindow({ forceRestore: false });
            } else {
                // Case 3: Window is visible and not minimized.
                if (isForeground) {
                    // Subcase 3a: It's in the foreground. Hide it.
                    this.log("🔄 Window is visible and foreground, hiding...");
                    this.hideMainWindow();
                } else {
                    // Subcase 3b: It's in the background. Bring it to the front.
                    this.log("🔄 Window is visible but background, bringing to front...");
                    this.bringToFront();
                }
            }
        } catch (e) {
            this.log(`✗ Error in handleTrayClick: ${e}`);
        }
    },

    hideMainWindow: function () {
        if (!this.getMainWindowHandle()) return;
        try {
            // This is the crucial part: we check and save the maximized state
            // *right before* we hide the window.
            this.windowWasMaximized = this.user32.IsZoomed(this.mainWindowHandle);
            this.log(`Hiding main window. Maximized state saved: ${this.windowWasMaximized}`);
            this.user32.ShowWindow(this.mainWindowHandle, this.constants.SW_HIDE);
            this.isWindowHidden = true;
        } catch (e) {
            this.log("✗ Error hiding main window: " + e);
        }
    },

    bringToFront: function () {
        if (!this.getMainWindowHandle()) {
            this.log('✗ No main window handle to bring to front.');
            return;
        }

        try {
            this.log('🖥️ Bringing window to front without changing state...');

            const hForegroundWnd = this.user32.GetForegroundWindow();
            const dwCurrentThreadId = this.kernel32.GetCurrentThreadId();
            const dwForegroundThreadId = this.user32.GetWindowThreadProcessId(hForegroundWnd, null);

            this.user32.AttachThreadInput(dwCurrentThreadId, dwForegroundThreadId, true);

            // Just set it as foreground. Don't use ShowWindow, as that could
            // change the maximized/restored state incorrectly.
            this.user32.SetForegroundWindow(this.mainWindowHandle);

            this.user32.AttachThreadInput(dwCurrentThreadId, dwForegroundThreadId, false);

            this.log('✓ Main window brought to front.');
        } catch (e) {
            this.log(`✗ Error bringing window to front: ${e}`);
        }
    },

    showMainWindow: function ({ forceRestore = false } = {}) {
        if (!this.getMainWindowHandle()) {
            this.log('✗ No main window handle to show.');
            return;
        }

        try {
            // If restoring from a minimized state (isIconic was true), we must use SW_RESTORE.
            // SW_RESTORE correctly restores a window to its previous state (maximized or normal).
            // Otherwise, restore based on the last saved value when we hid the window.
            const state = (forceRestore || !this.windowWasMaximized)
                ? this.constants.SW_RESTORE
                : this.constants.SW_MAXIMIZE;

            const stateName = state === this.constants.SW_MAXIMIZE ? 'Maximize' : 'Restore';
            this.log(`🖥️ Activating main window. ForceRestore=${forceRestore}. Final State: ${stateName} (${state})`);

            const hForegroundWnd = this.user32.GetForegroundWindow();
            const dwForegroundThreadId = this.user32.GetWindowThreadProcessId(hForegroundWnd, null);
            const dwCurrentThreadId = this.kernel32.GetCurrentThreadId();

            // Attach our thread's input processing to the foreground window's thread
            this.user32.AttachThreadInput(dwCurrentThreadId, dwForegroundThreadId, true);

            // Show the window in its correct state (maximized or restored)
            this.user32.ShowWindow(this.mainWindowHandle, state);
            this.user32.SetForegroundWindow(this.mainWindowHandle);

            // Detach the thread input
            this.user32.AttachThreadInput(dwCurrentThreadId, dwForegroundThreadId, false);

            this.isWindowHidden = false;
            this.log('✓ Main window shown.');

        } catch (e) {
            this.log(`✗ Error showing main window: ${e}`);
        }
    },

    cleanupHelper: function () {
        if (this.cleanupPerformed) return; // Prevent double-execution
        this.cleanupPerformed = true;

        this.isShuttingDown = true;

        // Force Kill using wscript + silent_kill.vbs (NO FLASHBANG)
        try {
            const sysProcess = this.Cc["@mozilla.org/process/util;1"].createInstance(this.Ci.nsIProcess);
            const sysFile = this.Cc["@mozilla.org/file/local;1"].createInstance(this.Ci.nsIFile);

            if (this.vbsPath) {
                // Use wscript.exe to run the VBS
                sysFile.initWithPath("C:\\Windows\\System32\\wscript.exe");
                sysProcess.init(sysFile);
                const args = [this.vbsPath];
                sysProcess.run(false, args, args.length);
                this.log("✓ Executed silent_kill.vbs");
            } else {
                // Fallback to noisy taskkill if VBS missing
                sysFile.initWithPath("C:\\Windows\\System32\\taskkill.exe");
                sysProcess.init(sysFile);
                const args = ["/F", "/IM", this.helperExeName];
                sysProcess.run(false, args, args.length);
            }
        } catch (e) {
            this.log("✗ Kill failed: " + e);
        }

        this.helperProcess = null;

        if (this.helperPath) {
            try {
                // Wait briefly for lock release
                const thread = this.Cc["@mozilla.org/thread-manager;1"].getService(this.Ci.nsIThreadManager).currentThread;
                // thread.dispatch(() => { /* delayed delete? */ }, 0);
            } catch (e) { }
        }
    },

    cleanup: function () {
        this.log("🧹 Cleaning up all resources...");
        this.isShuttingDown = true;

        if (this.hidePollingInterval) {
            clearInterval(this.hidePollingInterval);
            this.hidePollingInterval = null;
            this.log("✓ Polling interval cleared.");
        }

        if (this.prefPane) {
            Zotero.PreferencePanes.unregister(this.prefPane.paneID);
            this.prefPane = null;
            this.log("✓ Preferences pane unregistered.");
        }

        if (this.prefObserver) {
            Zotero.Prefs.unregisterObserver('extensions.zotero-in-tray.', this.prefObserver);
            this.log("✓ Preference observer unregistered.");
        }

        if (this.serverSocket) {
            this.serverSocket.close();
            this.log("✓ Server socket closed.");
        }

        this.cleanupHelper();

        for (let window of this.lockedWindows.keys()) {
            this.unlockWindow(window);
        }

        if (this.user32) this.user32.close();
        if (this.kernel32) this.kernel32.close();

        this.mainWindowHandle = null; // Clear handle on cleanup

        this.log("✓ Cleanup finished.");
    },

    tryHideWindowOnStartup: function () {
        if (this.initialHidePerformed || this.isShuttingDown) {
            if (this.hidePollingInterval) {
                clearInterval(this.hidePollingInterval);
                this.hidePollingInterval = null;
            }
            return;
        }

        // Only proceed if we have a valid main window object AND can get its handle
        // This prevents accidentally hiding Firefox or other Mozilla windows
        if (this.mainWindow && this.getMainWindowHandle()) {
            this.log('🚀 Zotero window handle is available. Hiding window now.');
            this.hideMainWindow();

            this.initialHidePerformed = true;
            clearInterval(this.hidePollingInterval);
            this.hidePollingInterval = null;
            this.log('✓ Initial auto-hide complete. Polling stopped.');
        } else {
            // SILENCED LOGGING to prevent UI Freeze/Disk I/O spam
            // this.log('⏳ Waiting for Zotero main window to be ready...');
        }
    }
};

// Global bootstrap functions
function install() { ZoteroInTray.log("Install event."); }
function uninstall() { ZoteroInTray.log("Uninstall event."); }
function startup({ id, version, rootURI }) {
    ZoteroInTray.init({ id, version, rootURI });
}
function shutdown() {
    ZoteroInTray.cleanup();
}
function onMainWindowLoad({ window }) {
    ZoteroInTray.log("🔥 Main window loaded: " + window.location.href);
    ZoteroInTray.mainWindow = window;
    ZoteroInTray.lockWindow(window);

    // Try to get window handle immediately
    if (!ZoteroInTray.getMainWindowHandle()) {
        ZoteroInTray.log("⏳ Initial handle acquisition failed, will retry...");
        // Retry after a short delay to ensure window is fully ready
        setTimeout(() => {
            if (!ZoteroInTray.getMainWindowHandle()) {
                ZoteroInTray.log("⚠️ Second attempt to get window handle failed");
                // Try one more time after document is fully loaded
                if (window.document.readyState !== 'complete') {
                    window.addEventListener('load', () => {
                        setTimeout(() => {
                            ZoteroInTray.getMainWindowHandle();
                        }, 100);
                    }, { once: true });
                }
            }
        }, 500);
    }
}
function onMainWindowUnload({ window }) {
    ZoteroInTray.unlockWindow(window);
    // Explicitly kill helper on window close to prevent Zombie Icon
    ZoteroInTray.cleanupHelper();

    if (ZoteroInTray.mainWindow === window) {
        ZoteroInTray.mainWindow = null;
        ZoteroInTray.mainWindowHandle = null;
    }
} 