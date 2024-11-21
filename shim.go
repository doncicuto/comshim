package comshim

import (
	"log"
	"runtime"
	"sync"

	"github.com/go-ole/go-ole"
)

// Shim provides control of a thread-locked goroutine that has been initialized
// for use with a mulithreaded component object model apartment. This is used
// to ensure that at least one thread within a process maintains an
// initialized connection to COM, and thus prevents COM resources from being
// unloaded from that process.
//
// Control is implemented through the use of a counter similar to a waitgroup.
// As long as the counter is greater than zero then the goroutine will remain
// in a blocked condition with its COM connection intact.
type Shim struct {
	m       sync.RWMutex
	cond    sync.Cond
	c       Counter // An atomic counter
	running bool
}

// New returns a new shim for keeping component object model resources allocated
// within a process.
func New() *Shim {
	shim := new(Shim)
	shim.cond.L = &shim.m
	return shim
}

// Add adds delta, which may be negative, to the counter for the shim. As long
// as the counter is greater than zero, at least one thread is guaranteed to be
// initialized for mutli-threaded COM access.
//
// If the counter becomes zero, the shim is released and COM resources may be
// released if there are no other threads that are still initialized.
//
// If the counter goes negative, Add panics.
//
// If the shim cannot be created for some reason, returns err
func (s *Shim) Add(delta int) error {
	// Check whether the shim is already running within a read lock
	s.m.RLock()
	if s.running {
		if err := s.add(delta); err != nil {
			return err
		}
		s.m.RUnlock()
		return nil
	}
	s.m.RUnlock()

	// The shim wasn't running; only change the running state within a write lock
	s.m.Lock()
	defer s.m.Unlock()
	if err := s.add(delta); err != nil {
		return err
	}
	if s.running {
		// The shim was started between the read lock and the write lock
		return nil
	}

	if err := s.run(); err != nil {
		return err
	}

	s.running = true
	return nil
}

// Done decrements the counter for the shim.
func (s *Shim) Done() error {
	return s.add(-1)
}

func (s *Shim) add(delta int) error {
	value := s.c.Add(int64(delta))
	if value == 0 {
		s.cond.Broadcast()
	}
	if value < 0 {
		log.Println("[ERROR]: found negative counter error")
		return ErrNegativeCounter
	}
	return nil
}

func (s *Shim) run() error {
	init := make(chan error)

	go func() {
		runtime.LockOSThread()
		defer runtime.UnlockOSThread()

		if err := ole.CoInitializeEx(0, ole.COINIT_MULTITHREADED); err != nil {
			switch err.(*ole.OleError).Code() {
			case 0x00000001: // S_FALSE
				// Some other goroutine called CoInitialize on this thread
				// before we ended up with it. This probably means the other
				// caller failed to lock the OS thread or failed to call
				// CoUninitialize.

				// We still decrement this thread's initialization counter by
				// calling CoUninitialize here, as recommended by the docs.
				ole.CoUninitialize()

				// Send an error so that shim.Add panics
				init <- ErrAlreadyInitialized
			default:
				init <- err
			}
			close(init)
			return
		}

		close(init)

		s.m.Lock()
		for s.c.Value() > 0 {
			s.cond.Wait()
		}
		s.running = false
		ole.CoUninitialize()
		s.m.Unlock()
	}()

	return <-init
}
