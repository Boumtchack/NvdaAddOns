import globalPluginHandler
from scriptHandler import script
import winUser
import api
from logHandler import log
import wx
import ctypes
from ctypes import wintypes

# Constants for the zoom levels. Don't put 1.0 as the minimum zoom level,
# as it will not disable the magnifier but just set it to 1x zoom.
ZOOM_MIN = 1.5
ZOOM_MAX = 10.0

# The step size for zooming in and out. This value is added or subtracted, keep the number a multiple of 5 to
# avoid floating point errors.
ZOOM_STEP = 0.5

# The interval in milliseconds for the timer that checks the review position.
# Adjust this value between 10 and 50 milliseconds for performance.
# A lower value means more frequent updates and a smoother mouse move, but may impact performance.
TIMER_INTERVAL_MS = 15

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    """NVDA Magnifier Global Plugin."""

    def __init__(self):
        super().__init__()
        self.timer = None
        self.magnifierIsOn = False
        self.zoom = 2.0
        self.lastMousePos = (0, 0)
        self.lastNVDAPos = (0, 0)
        self.lastX = 0
        self.lastY = 0 
        self._loadMagnification()

    @script(
        description="Toggle on and off the magnifier centering on the review position",
        category="Screen",
        gesture="kb:NVDA+shift+w"
    )
    def script_ToggleMagnifier(self, gesture):
        """
        Toggle the magnifier centering feature.
        If not enabled, enable it and start the timer to track focus.
        If enabled, disable it and reset the magnifier.
        """
        if not self.magnifierIsOn:
            self.magnifierIsOn = True
            self.timer = wx.CallLater(TIMER_INTERVAL_MS, self._chooseFocus)
            log.info(f"Magnifier centering enabled")
        else:
            self.magnifierIsOn = False
            self._resetMagnifier()
            log.info("Magnifier centering disabled")

    @script(
        description="Zoom in",
        category="Screen",
        gesture="kb:NVDA+shift+upArrow"
    )
    def script_zoomIn(self, gesture):
        """Increase magnifier zoom level."""
        self._zoom(+1)

    @script(
        description="Zoom out",
        category="Screen",
        gesture="kb:NVDA+shift+downArrow"
    )
    def script_zoomOut(self, gesture):
        """Decrease magnifier zoom level."""
        self._zoom(-1)

    def _zoom(self, direction):
        """
        Change the zoom level.
        direction: +1 to zoom in, -1 to zoom out.
        Only works if magnifier centering is enabled.
        """
        if self.magnifierIsOn:
            if direction > 0:
                self.zoom = min(self.zoom + ZOOM_STEP, ZOOM_MAX)
            else:
                self.zoom = max(self.zoom - ZOOM_STEP, ZOOM_MIN)
            log.info(f"Zoom level: {self.zoom}")
            self._centerMagnifier(self.lastX, self.lastY)
        else:
            log.info("Magnifier centering is not enabled at zoom change")

    def _loadMagnification(self):
        """
        Initialize the Magnification API.
        Tries to load the magnification DLL and initialize the API.
        """
        try:
            # Attempt to access the magnification DLL.
            ctypes.windll.magnification
        except (OSError, AttributeError):
            # If the DLL is not available, log this and exit the function.
            log.info("Magnification API not available")
            return
        # Try to initialize the magnification API.
        # MagInitialize returns 0 if already initialized or on failure.
        if ctypes.windll.magnification.MagInitialize() == 0:
            log.info("Magnification API already initialized")
            return
        # If initialization succeeded, log success.
        log.info("Magnification API initialized")

    def _getMagSetFullscreenTransform(self):
        """
        Helper to get the MagSetFullscreenTransform function from the magnification API,
        with correct return and argument types set.
        """
        # Get the MagSetFullscreenTransform function from the magnification API
        MagSetFullscreenTransform = ctypes.windll.magnification.MagSetFullscreenTransform
        # Set the return type of the function to BOOL (success/failure)
        MagSetFullscreenTransform.restype = wintypes.BOOL
        # Define the argument types: float for zoom, int for left, int for top
        MagSetFullscreenTransform.argtypes = [ctypes.c_float, ctypes.c_int, ctypes.c_int]
        return MagSetFullscreenTransform
    
    def _resetMagnifier(self):
        """
        Reset magnifier to default (1x zoom).
        Called when magnifier centering is disabled.
        """
        try:
            # Get the MagSetFullscreenTransform function from the magnification API (via helper)
            MagSetFullscreenTransform = self._getMagSetFullscreenTransform()
            # Call the function to reset the fullscreen magnifier:
            # - 1.0 for default zoom (no magnification)
            # - 0, 0 for the top-left corner of the screen
            result = MagSetFullscreenTransform(ctypes.c_float(1.0), ctypes.c_int(0), ctypes.c_int(0))
            # If the function call failed, log an error
            if not result:
                log.info("Failed to reset fullscreen transform")
            else:
                # If successful, log that the magnifier was reset
                log.info("Magnifier reset to 1x zoom")
        except AttributeError:
            # If the magnification API is not available, log an error
            log.info("Magnification API not available")
            
    def _getMagnifierWindow(self, x, y):
        """
        Compute the top-left corner of the magnifier window so that it is centered on (x, y).
        Returns (left, top).
        """
        zoom = self.zoom
        screenWidth = ctypes.windll.user32.GetSystemMetrics(0)
        screenHeight = ctypes.windll.user32.GetSystemMetrics(1)
        visibleWidth = screenWidth / zoom
        visibleHeight = screenHeight / zoom
        left = int(x - (visibleWidth / 2))
        top = int(y - (visibleHeight / 2))
        left = max(0, min(left, int(screenWidth - visibleWidth)))
        top = max(0, min(top, int(screenHeight - visibleHeight)))
        return (left, top)

    def _centerMagnifier(self, x, y):
        """
        Center the magnifier on the given (x, y) position.
        If the mouse moves, center on the mouse; otherwise, center on the NVDA review position.
        """
        zoom = self.zoom
        left, top = self._getMagnifierWindow(x, y)
        try:
            MagSetFullscreenTransform = self._getMagSetFullscreenTransform()
            result = MagSetFullscreenTransform(ctypes.c_float(zoom), ctypes.c_int(left), ctypes.c_int(top))
            if not result:
                log.info("Failed to set fullscreen transform")
        except AttributeError:
            log.info("Magnification API not available")
        self._continueMagnifier()
    
    def _continueMagnifier(self):
        """
        Continue or stop the magnifier timer.
        If magnifier centering is enabled, restart the timer.
        Otherwise, stop the timer and reset the magnifier.
        """
        if self.timer:
            self.timer.Stop()
        if self.magnifierIsOn:
            self.timer = wx.CallLater(TIMER_INTERVAL_MS, self._chooseFocus)
        else:
            self.timer = None
            self._resetMagnifier()
           
    def _getNVDAPosition(self):
        """
        Get the current review position as (x, y), falling back to navigator object if needed.
        Tries to get the review position from NVDA's API, or the center of the navigator object.
        """
        # Try to get the current review position object from NVDA's API
        reviewPosition = api.getReviewPosition()
        if reviewPosition:
            try:
                # Try to get the point at the start of the review position
                point = reviewPosition.pointAtStart
                return point.x, point.y
            except (NotImplementedError, LookupError, AttributeError):
                # If that fails, fall through to try navigator object
                pass

        # Fallback: try to use the navigator object location
        navigatorObject = api.getNavigatorObject()
        try:
            # Try to get the bounding rectangle of the navigator object
            left, top, width, height = navigatorObject.location
            # Calculate the center point of the rectangle
            x = left + (width // 2)
            y = top + (height // 2)
            return x, y
        except Exception:
            # If no location is found, log this and return (0, 0)
            log.info("No location found for navigator object.")
            return 0, 0
        
    def _isNVDAMoving(self):
        """
        Check if the NVDA review position has changed since the last check.
        """
        return self.lastNVDAPos != self._getNVDAPosition()
    
    def _getMousePosition(self):
        """
        Get the current mouse position as (x, y).
        """
        pt = wintypes.POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
        return (pt.x, pt.y)
    
    def _isMouseMoving(self):
        """
        Check if the mouse position has changed since the last check.
        """
        return self.lastMousePos != self._getMousePosition()
       
    def _chooseFocus(self):
        """
        Decide whether to center the magnifier on the NVDA review position or the mouse.
        If either has moved, update the last known position and center the magnifier.
        If neither has moved, continue or stop the magnifier as appropriate.
        """
        if self._isNVDAMoving():
            x, y = self._getNVDAPosition()
            self.lastNVDAPos = x, y
            self.lastX, self.lastY = x, y
            winUser.setCursorPos(x, y)
        
        elif self._isMouseMoving():
            x, y = self._getMousePosition()
            self.lastMousePos = x, y
            self.lastX, self.lastY = x, y
        
        else:
            self._continueMagnifier()
            return

        self._centerMagnifier(x, y)
