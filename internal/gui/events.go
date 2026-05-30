package gui

import appupdate "github.com/Fepozopo/bsc-hotsheet-update/internal/update"

type UIEvent interface {
	isUIEvent()
}

type generateCompletedEvent struct {
	Outputs []string
	Err     error
}

func (generateCompletedEvent) isUIEvent() {}

type updateCheckCompletedEvent struct {
	Result        appupdate.CheckResult
	Err           error
	ShowNoUpdates bool
}

func (updateCheckCompletedEvent) isUIEvent() {}

type selfUpdateCompletedEvent struct {
	Err error
}

func (selfUpdateCompletedEvent) isUIEvent() {}
