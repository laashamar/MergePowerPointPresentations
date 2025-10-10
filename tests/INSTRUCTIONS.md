

You are given a Python test suite for a PySide6 application located in the `tests` folder. Your task is to **analyze and automatically refactor the test code** to eliminate four common GUI testing anti‑patterns. Apply the following rules consistently across all test files:

1. **QApplication Management**  
   - Ensure there is a single, session‑scoped `qapp` fixture in `conftest.py`.  
   - Remove any manual instantiations of `QApplication([])` or `QApplication(sys.argv)` from test functions or module scope.  
   - Make all widget‑creating tests depend on the `qapp` fixture.

2. **Widget Interaction and Event Processing**  
   - For any test that interacts with widgets (clicks, typing, etc.), inject the `qtbot` fixture.  
   - Always call `qtbot.addWidget(widget)` and `widget.show()` before interaction.  
   - Replace direct calls like `button.click()` or `time.sleep()` with `qtbot.mouseClick()`, `qtbot.keyClicks()`, and `qtbot.waitUntil()`.

3. **Asynchronous Signal Handling**  
   - When a test triggers an action that emits a signal, wrap the interaction in `with qtbot.waitSignal(widget.signal, timeout=...)`.  
   - Move assertions after the signal has been received.  
   - If the signal carries arguments, use `blocker.args` to inspect them.

4. **Test Isolation**  
   - Ensure all fixtures that return GUI widgets are **function‑scoped** (default).  
   - Remove or refactor any module‑level widget instances so each test starts with a clean state.

### Deliverables
- Updated `conftest.py` with a correct `qapp` fixture.  
- Refactored test files where:  
  * Manual `QApplication` creation is removed.  
  * GUI interactions use `qtbot` properly.  
  * Asynchronous operations wait for signals.  
  * Fixtures are function‑scoped for isolation.  
- Preserve existing test logic and assertions, but make them robust and reproducible under pytest‑qt.