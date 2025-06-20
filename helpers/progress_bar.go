package helpers

import (
	"fmt"
)

type Bar struct {
	percent int64  // progress percentage
	cur     int64  // current progress
	total   int64  // total value for progress
	rate    string // the actual progress bar to be printed
	graph   string // the fill value for progress bar
}

// NewOption sets the start and total values for the progress bar,
// and initializes the progress bar with the given start value.
// If the graph character is empty, it defaults to '#'.
func (bar *Bar) NewOption(start, total int64) {
	bar.cur = start
	bar.total = total
	if bar.graph == "" {
		bar.graph = "#"
	}
	bar.percent = bar.getPercent()
	for i := 0; i < int(bar.percent); i += 2 {
		bar.rate += bar.graph // initial progress position
	}
}

// getPercent calculates the current progress percentage based on the current
// progress (cur) and the total value (total) of the progress bar. It returns
// the percentage as an int64 value.
func (bar *Bar) getPercent() int64 {
	return int64((float32(bar.cur) / float32(bar.total)) * 100)
}

// Play updates the progress bar with the current progress value (cur). It
// compares the new progress percentage with the previous one and adds a new
// graph character to the progress bar if the percentage has increased by 2% or
// more. It then prints the progress bar and the current progress value and
// total value.
func (bar *Bar) Play(cur int64) {
	bar.cur = cur
	bar.percent = bar.getPercent()
	barWidth := 50
	filled := int((float64(bar.percent) / 100.0) * float64(barWidth))
	if filled > barWidth {
		filled = barWidth
	}
	bar.rate = ""
	for i := 0; i < filled; i++ {
		bar.rate += bar.graph
	}
	for i := filled; i < barWidth; i++ {
		bar.rate += " "
	}
	fmt.Printf("\r[%-50s]%3d%% %8d/%d", bar.rate, bar.percent, bar.cur, bar.total)
}

// Finish prints a newline to the console to finish the progress bar.
func (bar *Bar) Finish() {
	fmt.Println()
}
