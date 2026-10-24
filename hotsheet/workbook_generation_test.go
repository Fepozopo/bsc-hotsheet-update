package hotsheet

import (
	"errors"
	"slices"
	"sync/atomic"
	"testing"
	"testing/synctest"
)

// generationTestGate returns a channel that releases a fake reader or writer.
// Cleanup releases blocked workers even if an assertion stops the test early.
func generationTestGate(t *testing.T) chan struct{} {
	t.Helper()
	gate := make(chan struct{})
	t.Cleanup(func() {
		select {
		case <-gate:
		default:
			close(gate)
		}
	})
	return gate
}

// TestWorkbookConcurrencyAndOrder verifies at most two writers overlap, freed
// capacity starts another line, and completion order does not reorder paths.
func TestWorkbookConcurrencyAndOrder(t *testing.T) {
	// synctest.Wait makes each blocked-writer checkpoint deterministic without
	// sleeps or timing assumptions, and synchronizes progress-slice inspection.
	synctest.Test(t, func(t *testing.T) {
		gates := map[string]chan struct{}{}
		for _, line := range []string{"A", "B", "C", "D"} {
			gates[line] = generationTestGate(t)
		}
		started := make(chan string, 4)
		done := make(chan struct{})
		var active atomic.Int32
		var outputs []string
		var buildErr error
		var progress []Progress
		go func() {
			outputs, buildErr = writeProductLineWorkbooks([]string{"A", "B", "C", "D"}, func(line string) (string, error) {
				if count := active.Add(1); count > 2 {
					t.Errorf("line %s: expected at most 2 active writers, got %d", line, count)
				}
				defer active.Add(-1)
				started <- line
				<-gates[line]
				return line + ".xlsx", nil
			}, func(update Progress) { progress = append(progress, update) })
			close(done)
		}()
		synctest.Wait()
		if len(started) != 2 {
			t.Fatalf("blocked initial writers: expected 2 starts, got %d", len(started))
		}
		initial := []string{<-started, <-started}
		slices.Sort(initial)
		if !slices.Equal(initial, []string{"A", "B"}) {
			t.Fatalf("expected initial writers A and B, got %v", initial)
		}
		for _, step := range []struct{ release, next string }{{"B", "C"}, {"C", "D"}} {
			close(gates[step.release])
			synctest.Wait()
			if len(started) != 1 {
				t.Fatalf("after %s completes: expected one replacement writer %s, got %d starts", step.release, step.next, len(started))
			}
			if got := <-started; got != step.next {
				t.Fatalf("after %s completes: expected writer %s, got %s", step.release, step.next, got)
			}
		}
		close(gates["D"])
		synctest.Wait()
		select {
		case <-done:
			t.Fatal("expected generation to wait for blocked writer A")
		default:
		}
		close(gates["A"])
		<-done
		want := []string{"A.xlsx", "B.xlsx", "C.xlsx", "D.xlsx"}
		if buildErr != nil || !slices.Equal(outputs, want) {
			t.Fatalf("completion order B,C,D,A: expected paths %v and no error, got %v and %v", want, outputs, buildErr)
		}
		previous := 0
		for _, update := range progress {
			if update.Percent < previous || update.Percent > 100 {
				t.Errorf("after %d%%: expected nondecreasing progress up to 100, got %+v", previous, update)
			}
			previous = update.Percent
		}
	})
}

// TestWorkbookFailureDrainsActiveWriters verifies scheduling stops on a failure,
// active work finishes before return, and its successful paths are not lost.
// If another active writer fails, the first observed error must be retained.
func TestWorkbookFailureDrainsActiveWriters(t *testing.T) {
	firstFailure := errors.New("first workbook failed")
	secondFailure := errors.New("second workbook failed")
	for _, tc := range []struct {
		name string
		err  error
		want []string
	}{
		{"active writer succeeds", nil, []string{"B.xlsx"}},
		{"active writer also fails", secondFailure, []string{}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			synctest.Test(t, func(t *testing.T) {
				a, b := generationTestGate(t), generationTestGate(t)
				started := make(chan string, 3)
				done := make(chan struct{})
				var outputs []string
				var buildErr error
				go func() {
					outputs, buildErr = writeProductLineWorkbooks([]string{"A", "B", "C"}, func(line string) (string, error) {
						started <- line
						switch line {
						case "A":
							<-a
							return "", firstFailure
						case "B":
							<-b
							return "B.xlsx", tc.err
						default:
							return "C.xlsx", nil
						}
					}, nil)
					close(done)
				}()
				synctest.Wait()
				close(a)
				synctest.Wait()
				if len(started) != 2 {
					t.Errorf("after A fails with B blocked: expected only 2 started lines, got %d", len(started))
				}
				select {
				case <-done:
					t.Fatal("returned after A failed while B was still writing")
				default:
				}
				close(b)
				<-done
				if !errors.Is(buildErr, firstFailure) || !slices.Equal(outputs, tc.want) {
					t.Errorf("A failed first and B returned %v: expected paths %v and error %v, got %v and %v", tc.err, tc.want, firstFailure, outputs, buildErr)
				}
			})
		})
	}
}

// TestWorkbookSchedulingBoundaries verifies empty and single-line inputs with
// nil progress callbacks, without starting unnecessary build calls.
func TestWorkbookSchedulingBoundaries(t *testing.T) {
	for _, tc := range []struct {
		name  string
		lines []string
		want  []string
	}{
		{"nil lines", nil, []string{}},
		{"empty lines", []string{}, []string{}},
		{"one line", []string{"ONLY"}, []string{"ONLY.xlsx"}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			calls := 0
			outputs, err := writeProductLineWorkbooks(tc.lines, func(line string) (string, error) {
				calls++
				return line + ".xlsx", nil
			}, nil)
			if err != nil || outputs == nil || !slices.Equal(outputs, tc.want) || calls != len(tc.want) {
				t.Errorf("lines %v: expected non-nil paths %v, %d builds, no error; got %v, %d builds, error %v", tc.lines, tc.want, len(tc.want), outputs, calls, err)
			}
		})
	}
}

// TestWorkbookFilenameCollisions verifies potentially identical destinations
// are rejected before any build can overwrite a competing product line's file.
func TestWorkbookFilenameCollisions(t *testing.T) {
	for _, tc := range []struct {
		name  string
		lines []string
	}{
		{"sanitized names", []string{"A:B", "AB"}},
		{"case insensitive names", []string{"BAS", "bas"}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			var called atomic.Bool
			outputs, err := writeProductLineWorkbooks(tc.lines, func(line string) (string, error) {
				called.Store(true)
				return line + ".xlsx", nil
			}, nil)
			if !errors.Is(err, ErrProductLineFilenameCollision) || len(outputs) != 0 || called.Load() {
				t.Errorf("colliding lines %v: expected filename-collision error without builds or outputs, got error %v, called=%v, outputs=%v", tc.lines, err, called.Load(), outputs)
			}
		})
	}
}
