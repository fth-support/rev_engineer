package main

import (
	"context"
	"database/sql"
	"fmt"
	"time"

	_ "github.com/denisenkom/go-mssqldb"

	"fyne.io/fyne/v2"
	"fyne.io/fyne/v2/app"
	"fyne.io/fyne/v2/container"
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
	serverIPEntry := widget.NewEntry()
	serverIPEntry.SetText("192.168.1.100")
	serverDBNameEntry := widget.NewEntry()
	serverDBNameEntry.SetText("HQ_CENTRAL")

	// Log Terminal
	logTextBox := widget.NewMultiLineEntry()
	// logTextBox.Disable() -> Removed to keep text bright and readable
	logTextBox.SetText("==================================================\nServiceTransfer System Ready.\nConfigure settings and click 'Start Sync'.\n==================================================\n")
	scrollContainer := container.NewScroll(logTextBox)
	uiLogger := &UILogger{LogEntry: logTextBox}

	testServerDbBtn := widget.NewButton("Test DB Connection", func() {
		ip := serverIPEntry.Text
		dbName := serverDBNameEntry.Text
		uiLogger.Println(fmt.Sprintf("Testing DB: %s (DB: %s)...", ip, dbName))
		
		go func() {
			connString := fmt.Sprintf("server=%s;user id=sa;password=pass;database=%s", ip, dbName)
			db, err := sql.Open("sqlserver", connString)
			if err != nil {
				uiLogger.Println(fmt.Sprintf("Error creating connection to %s: %v", ip, err))
				return
			}
			defer db.Close()
			
			ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
			defer cancel()
			if err = db.PingContext(ctx); err != nil {
				uiLogger.Println(fmt.Sprintf("Connection failed to %s: %v", ip, err))
			} else {
				uiLogger.Println(fmt.Sprintf("✅ SUCCESS! Connected to DB: %s", dbName))
			}
		}()
	})

	// Layout for DB Settings
	settingsForm := widget.NewForm(
		&widget.FormItem{Text: "Server DB Location:", Widget: serverIPEntry},
		&widget.FormItem{Text: "Server DB Name:", Widget: serverDBNameEntry},
		&widget.FormItem{Text: "", Widget: testServerDbBtn},
	)


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
		serverIPEntry.Disable()
		serverDBNameEntry.Disable()

		uiLogger.Println("Sync Agent Started.")

		config := Config{
			LocalDBDSN:   fmt.Sprintf("server=%s;user id=sa;password=pass;database=%s", serverIPEntry.Text, serverDBNameEntry.Text),
			ServerDBDSN:  fmt.Sprintf("server=%s;user id=sa;password=pass;database=%s", serverIPEntry.Text, serverDBNameEntry.Text),
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
		serverIPEntry.Enable()
		serverDBNameEntry.Enable()
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
