import globalPluginHandler
from scriptHandler import script
import winUser
import api
from logHandler import log
import wx
import ctypes
from ctypes import wintypes
import ui

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
TIMER_INTERVAL_MS = 10
MARGIN_BORDER = 50

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
    """NVDA Magnifier Global Plugin."""

    def __init__(self):
        super().__init__()
        self.timer = None
        self.magnifier_is_on = False
        self.zoom = 2.0
        self.last_mouse_pos = (0, 0)
        self.last_nvda_pos = (0, 0)
        self.last_screen_left = 0
        self.last_screen_top = 0
        self.mode = "center"
        self._loadMagnification()

    @script(
        description="Toggle on and off the magnifier centering on the review position",
        category="Screen",
        gesture="kb:NVDA+shift+w"
    )
    def script_toggleMagnifier(self, gesture):
        """
        Toggle the magnifier centering feature.
        If not enabled, enable it and start the timer to track focus.
        If enabled, disable it and reset the magnifier.
        """
        if not self.magnifier_is_on:
            self.magnifier_is_on = True
            self.timer = wx.CallLater(TIMER_INTERVAL_MS, self._chooseFocus)
            # english
            # ui.message(f"Magnifier started with mode {self.mode}")
            # français
            ui.message(f"la loupe a débuté avec le mode {'centré' if self.mode == 'center' else 'bordure'}")
        else:
            self.magnifier_is_on = False
            self._resetMagnifier()       
            # english
            # ui.message("Magnifier stoped")
            # français
            ui.message("La loupe s'arrête")

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

    @script(
        description="toggle mouse between center and border",
        category="Screen",
        gesture="kb:NVDA+shift+q"
    )
    def script_toggleMouseMode(self, gesture):
        if self.magnifier_is_on:
            if self.mode == "center":
                self.mode = "border"
                # english
                # ui.message("mode changed to border")
                # français
                ui.message("mode de zoom changé à bordure")
            else:
                self.mode = "center"
                # english
                # ui.message("mode changed to center")
                # français
                ui.message("mode de zoom changé à centré")
        else:
            ui.message("activez la loupe avec NVDA maj w avant de changer de mode")

    def _zoom(self, direction):
        """
        Change the zoom level.
        direction: +1 to zoom in, -1 to zoom out.
        Only works if magnifier centering is enabled.
        """
        if self.magnifier_is_on:
            if direction > 0:
                self.zoom = min(self.zoom + ZOOM_STEP, ZOOM_MAX)
            else:
                self.zoom = max(self.zoom - ZOOM_STEP, ZOOM_MIN)
            # english
            # ui.message(f"Zoom level: {self.zoom}")
            # français
            ui.message(f"Niveau de zoom changé à {self.zoom}")
            self._centerMagnifier(self.last_screen_left, self.last_screen_top)
        else:
            # english
            # ui.message(f"Niveau de zoom changé à {self.zoom}")
            # français
            ui.message("activez la loupe avec NVDA maj w avant de zoomer")

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
            # 1.0 for default zoom (no magnification)
            # 0, 0 for the top-left corner of the screen
            result = MagSetFullscreenTransform(ctypes.c_float(1.0), ctypes.c_int(0), ctypes.c_int(0))
            # If the function call failed, log an error
            if not result:
                log.info("Failed to reset fullscreen transform")
            else:
                log.info("Magnifier reset to 1x zoom")
        except AttributeError:
            # If the magnification API is not available, log an error
            log.info("Magnification API not available")

    def _getMagnifierWindow(self, x, y):
        """
        Compute the top-left corner of the magnifier window so that it is centered on (x, y).
        Returns (left, top, visible_width, visible_height).
        - x, y: The coordinates (in screen pixels) where you want the magnifier to be centered.
        - It then computes the top-left corner (left, top) so that the visible area is centered on (x, y).
        - If centering would make the window go off-screen, it clamps (left, top) so the visible area stays fully on screen.
        """
        # Get the screen size in pixels
        screen_width, screen_height = ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1)
        
        # Calculate the size of the visible area at the current zoom level
        visible_width, visible_height = screen_width / self.zoom, screen_height / self.zoom 
        
        # Compute the top-left corner so that (x, y) is at the center of the visible area
        left, top = int(x - (visible_width / 2)), int(y - (visible_height / 2))
        
        # Clamp the top-left corner so the visible area stays within the screen boundaries
        left, top = max(0, min(left, int(screen_width - visible_width))), max(0, min(top, int(screen_height - visible_height)))
        
        return (left, top, visible_width, visible_height)

    def _centerMagnifier(self, x, y):
        """
        Center the magnifier on the given (x, y) position.
        If the mouse moves, center on the mouse; otherwise, center on the NVDA review position.
        """
        left, top, visible_width, visible_height = self._getMagnifierWindow(x, y)
        try:
            MagSetFullscreenTransform = self._getMagSetFullscreenTransform()
            result = MagSetFullscreenTransform(ctypes.c_float(self.zoom), ctypes.c_int(left), ctypes.c_int(top))
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
        if self.magnifier_is_on:
            self.timer = wx.CallLater(TIMER_INTERVAL_MS, self._chooseFocus)
        else:
            self.timer = None
            self._resetMagnifier()

    def _getNVDAPosition(self):
        """
        Get the current review position as (x, y), falling back to navigator object if needed.
        Tries to get the review position from NVDA's API, or the center of the navigator object.
        This part is taken from NVDA+shift+m gesture.
        """
        # Try to get the current review position object from NVDA's API
        review_position = api.getReviewPosition()
        if review_position:
            try:
                # Try to get the point at the start of the review position
                point = review_position.pointAtStart
                return point.x, point.y
            except (NotImplementedError, LookupError, AttributeError):
                # If that fails, fall through to try navigator object
                pass

        # Fallback: try to use the navigator object location
        navigator_object = api.getNavigatorObject()
        try:
            # Try to get the bounding rectangle of the navigator object
            left, top, width, height = navigator_object.location
            # Calculate the center point of the rectangle
            x = left + (width // 2)
            y = top + (height // 2)
            return x, y
        except Exception:
            # If no location is found, log this and return (0, 0)
            log.info("No location found for navigator object.")
            return 0, 0

    def _getMousePosition(self):
        """
        Get the current mouse position as (x, y).
        """
        pt = wintypes.POINT()
        ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
        return (pt.x, pt.y)

    def _isMouseNearBorder(self):
        """
        Check if the mouse is near the border of the magnifier window.
        If so, adjust x, y to keep the mouse just at the margin limit.
        Otherwise, return last_x and last_y.
        """
        left, top, visible_width, visible_height = self._getMagnifierWindow(self.last_screen_left, self.last_screen_top)
        mouse_x, mouse_y = self._getMousePosition()
        self.last_mouse_pos = mouse_x, mouse_y

        min_x = left + MARGIN_BORDER
        max_x = left + visible_width - MARGIN_BORDER
        min_y = top + MARGIN_BORDER
        max_y = top + visible_height - MARGIN_BORDER

        dx = 0
        dy = 0

        if mouse_x < min_x:
            dx = mouse_x - min_x
        elif mouse_x > max_x:
            dx = mouse_x - max_x

        if mouse_y < min_y:
            dy = mouse_y - min_y
        elif mouse_y > max_y:
            dy = mouse_y - max_y

        if dx != 0 or dy != 0:
            self.last_screen_left += dx
            self.last_screen_top += dy

        return self.last_screen_left, self.last_screen_top

    # def _chooseFocus(self):
    #     """
    #     Decide whether to center the magnifier on the NVDA review position or the mouse.
    #     If either has moved, update the last known position and center the magnifier.
    #     If neither has moved, continue or stop the magnifier as appropriate.
    #     """
    #     nvda_pos = self._getNVDAPosition()
    #     mouse_pos = self._getMousePosition()

    #     if self.last_nvda_pos != nvda_pos:
    #         x, y = nvda_pos
    #         self.last_nvda_pos = self.last_screen_left, self.last_screen_top = x, y
    #         winUser.setCursorPos(x, y)
    #     elif self.last_mouse_pos != mouse_pos:
    #         if self.mode == "border":
    #             x, y = self._isMouseNearBorder()
    #         else:
    #             x, y = mouse_pos
    #             self.last_mouse_pos = self.last_screen_left, self.last_screen_top = x, y
    #     else:
    #         self._continueMagnifier()
    #         return

    #     self._centerMagnifier(x, y)

    # Séparation en deux pour plus de lisibilité

    def _focusOnNvda(self, pos):
        self.last_nvda_pos = pos
        self.last_screen_left, self.last_screen_top = pos
        winUser.setCursorPos(*pos)
        return pos

    def _focusOnMouse(self, pos):
        self.last_mouse_pos = pos
        if self.mode == "border":
            return self._isMouseNearBorder()
        self.last_screen_left, self.last_screen_top = pos
        return pos


    def _chooseFocus(self):
        nvda_pos = self._getNVDAPosition()
        mouse_pos = self._getMousePosition()

        if self.last_nvda_pos != nvda_pos:
            focus_pos = self._focusOnNvda(nvda_pos)
        elif self.last_mouse_pos != mouse_pos:
            focus_pos = self._focusOnMouse(mouse_pos)
        else:
            self._continueMagnifier()
            return

        self._centerMagnifier(*focus_pos)
