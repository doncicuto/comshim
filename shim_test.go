package comshim_test

import (
	"log"
	"runtime"
	"sync"
	"testing"

	"github.com/doncicuto/comshim"
	"github.com/go-ole/go-ole/oleutil"
)

func TestConcurrentShims(t *testing.T) {
	var maxRounds int
	if testing.Short() {
		maxRounds = 64
	} else {
		maxRounds = 256
	}

	// Vary the number of threads
	for procs := 1; procs < 11; procs++ {
		runtime.GOMAXPROCS(procs)

		// Vary the number of shims
		for rounds := 1; rounds <= maxRounds; rounds *= 2 {
			wg := sync.WaitGroup{}
			for i := 0; i < rounds; i++ {
				wg.Add(1)
				go func(i int) {
					defer wg.Done()

					if err := comshim.Add(1); err != nil {
						log.Printf("[ERROR]: found error adding a thread: %v", err)
						return
					}
					defer func() {
						if err := comshim.Done(); err != nil {
							log.Printf("[ERROR]: found error calling Done: %v", err)
						}
					}()

					obj, err := oleutil.CreateObject("WbemScripting.SWbemLocator")
					if err != nil {
						t.Error(err)
					} else {
						defer obj.Release()
					}
				}(i)
			}
			wg.Wait()
		}
	}
}
