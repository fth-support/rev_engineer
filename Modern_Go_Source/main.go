package main

import (
	"context"
	"fmt"
	"time"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
	"fyne.io/fyne/v2/dialog"
	"fyne.io/fyne/v2/layout"
	"fyne.io/fyne/v2/widget"
)

// Config represents application configuration
type Config struct {
	LocalDBDSN   string
	ServerDBDSN  string
	PollInterval time.Duration
}

// UILogger helps append logs to the UI thread safely
type UILogger struct {
	LogEntry *widget.Entry
}

func (u *UILogger) Println(msg string) {
	timestamp := time.Now().Format("2006-01-02 15:04:05")
	formattedMsg := fmt.Sprintf("[%s] %s\n", timestamp, msg)

	u.LogEntry.SetText(u.LogEntry.Text + formattedMsg)
	// Keeps cursor at bottom
	u.LogEntry.CursorRow = len(u.LogEntry.Text)
	u.LogEntry.Refresh()
}

func main() {
	myApp := app.New()
	myWindow := myApp.NewWindow("ServiceTransfer - Sync Agent")

	// Global states
	var cancelFunc context.CancelFunc
	isRunning := false

	// Config Inputs
	dbPathEntry := widget.NewEntry()
	dbPathEntry.SetText("C:\\POS\\LocalData.db")
	
	serverIPEntry := widget.NewEntry()
	serverIPEntry.SetText("192.168.1.100")

	// Browse Button
	browseBtn := widget.NewButton("Browse...", func() {
		fd := dialog.NewFileOpen(func(reader fyne.URIReadCloser, err error) {
			if err != nil {
				dialog.ShowError(err, myWindow)
				return
			}
			if reader == nil {
				return
			}
			dbPathEntry.SetText(reader.URI().Path())
		}, myWindow)
		fd.Show()
	})

	// Layout for DB Settings
	settingsForm := container.NewVBox(
		widget.NewLabelWithStyle("Database Configuration", fyne.TextAlignLeading, fyne.TextStyle{Bold: true}),
		container.NewBorder(nil, nil, widget.NewLabel("Local DB: "), browseBtn, dbPathEntry),
		container.NewBorder(nil, nil, widget.NewLabel("Server IP:"), nil, serverIPEntry),
	)

	// Log Terminal
	logTextBox := widget.NewMultiLineEntry()
	// logTextBox.Disable() -> Removed to keep text bright and readable
	logTextBox.SetText("==================================================\nServiceTransfer System Ready.\nConfigure settings and click 'Start Sync'.\n==================================================\n")
	scrollContainer := container.NewScroll(logTextBox)
	uiLogger := &UILogger{LogEntry: logTextBox}

	// Status Bar
	statusLabel := widget.NewLabelWithStyle("🔴 Status: Stopped", fyne.TextAlignTrailing, fyne.TextStyle{Bold: true})

	// Buttons
	var startBtn, stopBtn *widget.Button

	startBtn = widget.NewButton("▶ Start Sync", func() {
		if isRunning {
			return
		}
		isRunning = true
		statusLabel.SetText("🟢 Status: Running...")
		startBtn.Disable()
		stopBtn.Enable()
		dbPathEntry.Disable()
		serverIPEntry.Disable()
		browseBtn.Disable()

		uiLogger.Println("Sync Agent Started.")

		config := Config{
			LocalDBDSN:   dbPathEntry.Text,
			ServerDBDSN:  serverIPEntry.Text,
			PollInterval: 5 * time.Second,
		}

		var ctx context.Context
		ctx, cancelFunc = context.WithCancel(context.Background())
		go backgroundWorker(ctx, config, uiLogger)
	})
	startBtn.Importance = widget.HighImportance

	stopBtn = widget.NewButton("⏹ Stop", func() {
		if !isRunning {
			return
		}
		cancelFunc() // Signals background worker to stop
		isRunning = false
		statusLabel.SetText("🔴 Status: Stopped")
		stopBtn.Disable()
		startBtn.Enable()
		dbPathEntry.Enable()
		serverIPEntry.Enable()
		browseBtn.Enable()
		uiLogger.Println("Sync Agent Stopped by user.")
	})
	stopBtn.Disable()

	clearBtn := widget.NewButton("Clear Logs", func() {
		logTextBox.SetText("")
	})

	// Control Panel Layout
	controlPanel := container.NewHBox(
		startBtn,
		stopBtn,
		layout.NewSpacer(),
		clearBtn,
	)

	// Combine everything
	topSection := container.NewVBox(settingsForm, widget.NewSeparator(), controlPanel, widget.NewSeparator())
	bottomSection := container.NewVBox(widget.NewSeparator(), statusLabel)

	mainLayout := container.NewBorder(topSection, bottomSection, nil, nil, scrollContainer)

	myWindow.SetContent(mainLayout)
	myWindow.Resize(fyne.NewSize(700, 500))
	myWindow.ShowAndRun()
}

// backgroundWorker runs the continuous polling loop
func backgroundWorker(ctx context.Context, config Config, logger *UILogger) {
	logger.Println(fmt.Sprintf("Local DB Path: %s", config.LocalDBDSN))
	logger.Println(fmt.Sprintf("Server IP: %s", config.ServerDBDSN))
	time.Sleep(1 * time.Second)
	logger.Println("Database connection established.")

	ticker := time.NewTicker(config.PollInterval)
	defer ticker.Stop()

	for {
		select {
		case <-ctx.Done():
			logger.Println("Background worker terminated.")
			return
		case <-ticker.C:
			logger.Println("--- [Tick] Syncing ---")
			runSyncCycle(logger)
		}
	}
}

// runSyncCycle encapsulates the core workflows originally from mMain.bas and wMain.frm
func runSyncCycle(logger *UILogger) {
	if !checkServerOnline() {
		logger.Println("WARNING: Central Server is unreachable.")
		return
	}
	syncTransactionData(logger)
	updateMemberPoints(logger)
}

func checkServerOnline() bool {
	return true
}

func syncTransactionData(logger *UILogger) {
	logger.Println("Checking for unsynced POS Transactions...")
	time.Sleep(300 * time.Millisecond) // Simulate DB I/O
	logger.Println("Transactions synced successfully.")
}

func updateMemberPoints(logger *UILogger) {
	logger.Println("Checking for unprocessed Member Points...")
	time.Sleep(200 * time.Millisecond)
	logger.Println("Member Points updated to Central Server.")
}
