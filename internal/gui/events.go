package gui

import appupdate "github.com/Fepozopo/bsc-hotsheet-update/internal/update"

// UIEvent is the marker interface used for messages sent from background
// goroutines back to the immediate-mode GUI state.
//
// The event channel keeps long-running work off the UI thread while still
// applying results on the next render pass.
type UIEvent interface {
	isUIEvent()
}

// generateCompletedEvent reports the outcome of a background hotsheet
// generation run.
type generateCompletedEvent struct {
	Outputs []string
	Err     error
}

// isUIEvent is a marker method to satisfy the UIEvent interface.
func (generateCompletedEvent) isUIEvent() {}

// updateCheckCompletedEvent reports the result of a startup or manual update
// check.
type updateCheckCompletedEvent struct {
	Result        appupdate.CheckResult
	Err           error
	ShowNoUpdates bool
}

func (updateCheckCompletedEvent) isUIEvent() {}

// selfUpdateCompletedEvent reports the outcome of applying a downloaded update.
type selfUpdateCompletedEvent struct {
	Err error
}

func (selfUpdateCompletedEvent) isUIEvent() {}
