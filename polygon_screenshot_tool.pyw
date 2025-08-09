# Save this file with a .pyw extension (e.g., polygon_screenshot_service.pyw) to run without a console window.
import tkinter as tk
from tkinter import messagebox
import math
import ctypes
from ctypes import wintypes
import sys
import os
import threading
import time

# Lazy import optimization - only import heavy modules when needed
PIL_IMPORTED = False
WIN32_IMPORTED = False
KEYBOARD_IMPORTED = False

def lazy_import_pil():
    """Import PIL modules only when needed"""
    global PIL_IMPORTED, Image, ImageGrab, ImageDraw
    if not PIL_IMPORTED:
        try:
            from PIL import Image, ImageGrab, ImageDraw
            PIL_IMPORTED = True
        except ImportError:
            messagebox.showerror("Missing Dependency", 
                               "PIL/Pillow is required. Install with: pip install Pillow")
            sys.exit(1)

def lazy_import_win32():
    """Import win32 modules only when needed"""
    global WIN32_IMPORTED, win32clipboard, io
    if not WIN32_IMPORTED:
        try:
            import win32clipboard
            import io
            WIN32_IMPORTED = True
            return win32clipboard, io
        except ImportError:
            return None, None
    return win32clipboard, io

def lazy_import_keyboard():
    """Import keyboard module only when needed"""
    global KEYBOARD_IMPORTED, keyboard
    if not KEYBOARD_IMPORTED:
        try:
            import keyboard
            KEYBOARD_IMPORTED = True
            return keyboard
        except ImportError:
            messagebox.showerror("Missing Dependency", 
                               "keyboard module is required. Install with: pip install keyboard")
            sys.exit(1)
    return keyboard

class SystemTrayIcon:
    """System tray icon implementation"""
    def __init__(self, screenshot_tool):
        self.screenshot_tool = screenshot_tool
        self.root = tk.Tk()
        self.root.withdraw()  # Start hidden
        
        # Create system tray menu
        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label="Capture Screenshot (Ctrl+Alt+2)", command=self.trigger_capture)
        self.menu.add_separator()
        self.menu.add_command(label="Show Window", command=self.show_window)
        self.menu.add_command(label="Exit", command=self.exit_app)
        
        # Create a hidden window for tray functionality
        self.tray_window = tk.Toplevel(self.root)
        self.tray_window.withdraw()
        self.tray_window.title("Polygon Screenshot Tool")
        
        # Set up the tray icon using Windows API
        self._setup_tray_icon()
        
        # Bind right-click to show menu
        self.root.bind("<Button-3>", self.show_menu)

    def _setup_tray_icon(self):
        """Setup system tray icon using Windows API"""
        try:
            # This is a simplified approach - in a full implementation,
            # you'd use win32gui and win32api for proper system tray
            self.root.iconbitmap(default=None)  # Use default icon
        except:
            pass

    def show_menu(self, event=None):
        """Show the context menu"""
        try:
            self.menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.menu.grab_release()

    def show_window(self):
        """Show a simple status window"""
        status_window = tk.Toplevel(self.root)
        status_window.title("Polygon Screenshot Tool")
        status_window.geometry("300x150")
        status_window.attributes('-topmost', True)
        
        # Center the window
        status_window.update_idletasks()
        x = (status_window.winfo_screenwidth() // 2) - (300 // 2)
        y = (status_window.winfo_screenheight() // 2) - (150 // 2)
        status_window.geometry(f"+{x}+{y}")
        
        tk.Label(status_window, text="Polygon Screenshot Tool", 
                font=("Arial", 14, "bold")).pack(pady=10)
        tk.Label(status_window, text="Press Ctrl+Alt+2 to capture", 
                font=("Arial", 10)).pack(pady=5)
        tk.Label(status_window, text="Running in system tray", 
                font=("Arial", 10), fg="green").pack(pady=5)
        
        button_frame = tk.Frame(status_window)
        button_frame.pack(pady=20)
        
        tk.Button(button_frame, text="Capture Now", 
                 command=self.trigger_capture, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Close", 
                 command=status_window.destroy, width=8).pack(side=tk.LEFT, padx=5)

    def trigger_capture(self):
        """Trigger screenshot capture"""
        threading.Thread(target=self.screenshot_tool.start_capture, daemon=True).start()

    def exit_app(self):
        """Exit the application"""
        self.screenshot_tool.stop()
        self.root.quit()

class FastPolygonScreenshotTool:
    """
    Background polygon screenshot tool with global hotkey support
    """
    PROXIMITY_THRESHOLD = 15

    def __init__(self):
        self.active = False
        self.running = True
        
        # System metrics cache
        self.dpi_scale = self._get_dpi_scale_fast()
        self._cache_system_metrics()
        
        # Initialize UI components as None
        self.root = None
        self.top_level = None
        self.canvas = None
        
        # State
        self.polygon_points = []
        self.is_closed = False

    def _get_dpi_scale_fast(self):
        """Fast DPI detection with proper high-DPI handling"""
        try:
            # Make the process DPI aware
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
            
            hdc = ctypes.windll.user32.GetDC(0)
            dpi_x = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88)  # LOGPIXELSX
            dpi_y = ctypes.windll.gdi32.GetDeviceCaps(hdc, 90)  # LOGPIXELSY
            ctypes.windll.user32.ReleaseDC(0, hdc)
            
            scale_x = dpi_x / 96.0 if dpi_x > 0 else 1.0
            scale_y = dpi_y / 96.0 if dpi_y > 0 else 1.0
            
            print(f"Debug: DPI - X: {dpi_x}, Y: {dpi_y}, Scale: {scale_x}x{scale_y}")
            return max(scale_x, scale_y)  # Use the larger scale factor
        except:
            return 1.0

    def _cache_system_metrics(self):
        """Cache all system metrics with proper DPI handling"""
        # Set DPI awareness first
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except:
            pass
            
        SM_XVIRTUALSCREEN = 76
        SM_YVIRTUALSCREEN = 77
        SM_CXVIRTUALSCREEN = 78
        SM_CYVIRTUALSCREEN = 79
        
        GetSystemMetrics = ctypes.windll.user32.GetSystemMetrics
        
        # Get logical coordinates (what Windows reports)
        self.virtual_x = GetSystemMetrics(SM_XVIRTUALSCREEN)
        self.virtual_y = GetSystemMetrics(SM_YVIRTUALSCREEN)
        self.virtual_width = GetSystemMetrics(SM_CXVIRTUALSCREEN)
        self.virtual_height = GetSystemMetrics(SM_CYVIRTUALSCREEN)
        
        # Get primary monitor size
        self.primary_width = GetSystemMetrics(0)  # SM_CXSCREEN
        self.primary_height = GetSystemMetrics(1)  # SM_CYSCREEN
        
        print(f"Debug: Logical screen: {self.virtual_x}, {self.virtual_y}, {self.virtual_width}x{self.virtual_height}")
        print(f"Debug: Primary monitor logical: {self.primary_width}x{self.primary_height}")
        print(f"Debug: Your laptop resolution: 2560x1600")
        print(f"Debug: DPI Scale factor: {self.dpi_scale}")
        
        # Calculate actual physical coordinates if needed
        self.physical_width = int(self.virtual_width * self.dpi_scale)
        self.physical_height = int(self.virtual_height * self.dpi_scale)
        print(f"Debug: Calculated physical: {self.physical_width}x{self.physical_height}")

    def start_capture(self):
        """Start capture mode"""
        if self.active:
            return
            
        self.active = True
        self.polygon_points = []
        self.is_closed = False
        
        # Create UI for capture
        self._create_capture_ui()

    def _create_capture_ui(self):
        """Create the capture UI"""
        # Create new root for capture session
        self.root = tk.Tk()
        self.root.withdraw()
        
        self.top_level = tk.Toplevel(self.root)
        
        # Setup window
        attrs = {
            '-topmost': True,
            '-alpha': 0.3
        }
        for attr, value in attrs.items():
            self.top_level.attributes(attr, value)
        
        self.top_level.configure(bg='black')
        self.top_level.overrideredirect(True)
        self.top_level.geometry(f"{self.virtual_width}x{self.virtual_height}+{self.virtual_x}+{self.virtual_y}")
        
        # Canvas setup
        self.canvas = tk.Canvas(
            self.top_level,
            width=self.virtual_width,
            height=self.virtual_height,
            bg='black',
            highlightthickness=0,
            cursor='crosshair'
        )
        self.canvas.pack()
        
        # Bind events
        self._bind_events()
        
        # Show window
        self.top_level.deiconify()
        self.top_level.focus_force()
        
        # Start the capture loop
        self.root.mainloop()

    def _bind_events(self):
        """Bind all events"""
        bindings = [
            (self.canvas, "<Button-1>", self.add_point),
            (self.canvas, "<Motion>", self.on_motion),
            (self.canvas, "<Button-3>", self.close_polygon),
            (self.canvas, "<Double-Button-1>", self.close_polygon),
            (self.top_level, "<Return>", self.capture_polygon),
            (self.top_level, "<Escape>", self.cancel_capture),
            (self.top_level, "<c>", self.clear_all),
            (self.top_level, "<BackSpace>", self.undo_point),
        ]
        
        for widget, event, handler in bindings:
            widget.bind(event, handler)

    def canvas_to_screen_coords(self, canvas_x, canvas_y):
        """Convert canvas coordinates to screen coordinates"""
        return canvas_x + self.virtual_x, canvas_y + self.virtual_y

    def add_point(self, event):
        """Add point with proximity check"""
        if self.is_closed:
            return

        x, y = event.x, event.y

        # Auto-close check
        if len(self.polygon_points) > 2:
            start_x, start_y = self.polygon_points[0]
            if math.hypot(start_x - x, start_y - y) < self.PROXIMITY_THRESHOLD:
                self.close_polygon(event)
                return

        self.polygon_points.append((x, y))
        self.draw_polygon_fast()

    def on_motion(self, event):
        """Motion handler with optimized drawing"""
        if not self.polygon_points or self.is_closed:
            return
        
        self.canvas.delete("preview_line", "close_indicator")
        
        last_x, last_y = self.polygon_points[-1]
        self.canvas.create_line(
            last_x, last_y, event.x, event.y,
            fill="yellow", width=1, dash=(4, 4),
            tags="preview_line"
        )

        if len(self.polygon_points) > 2:
            start_x, start_y = self.polygon_points[0]
            if math.hypot(start_x - event.x, start_y - event.y) < self.PROXIMITY_THRESHOLD:
                self._draw_close_indicator(start_x, start_y)

    def _draw_close_indicator(self, x, y):
        """Draw close indicator"""
        t = self.PROXIMITY_THRESHOLD
        self.canvas.create_oval(
            x - t, y - t, x + t, y + t,
            outline="red", width=3, fill="yellow",
            tags="close_indicator"
        )
        self.canvas.create_text(
            x, y - t - 20, text="CLICK TO CLOSE",
            fill="red", font=("Arial", 10, "bold"),
            tags="close_indicator"
        )

    def draw_polygon_fast(self, is_final=False):
        """Optimized polygon drawing"""
        self.canvas.delete("polygon")
        if not self.polygon_points:
            return

        if len(self.polygon_points) > 1:
            self.canvas.create_line(
                self.polygon_points, fill="lime", width=2, tags="polygon"
            )

        for i, (x, y) in enumerate(self.polygon_points):
            color = "yellow" if i == 0 else "lime"
            self.canvas.create_oval(
                x - 3, y - 3, x + 3, y + 3,
                fill=color, outline="white", tags="polygon"
            )

        if is_final and len(self.polygon_points) > 2:
            flat_points = [coord for point in self.polygon_points for coord in point]
            self.canvas.create_polygon(
                flat_points, fill="red", stipple="gray25", tags="polygon"
            )

    def close_polygon(self, event):
        """Close polygon and auto-capture"""
        if len(self.polygon_points) < 3 or self.is_closed:
            return

        self.is_closed = True
        self.canvas.delete("preview_line", "close_indicator")
        self.draw_polygon_fast(is_final=True)
        
        # Hide the overlay immediately
        self.top_level.withdraw()
        # REDUCED DELAY - much faster capture
        self.root.after(50, self._perform_capture)  # Changed from 300ms to 50ms

    def undo_point(self, event):
        """Remove last point"""
        if self.polygon_points and not self.is_closed:
            self.polygon_points.pop()
            self.draw_polygon_fast()

    def capture_polygon(self, event):
        """Manual capture trigger"""
        if not self.is_closed or len(self.polygon_points) < 3:
            return
        self.top_level.withdraw()
        self.root.after(50, self._perform_capture)  # Changed from 100ms to 50ms

    def _perform_capture(self):
        """Optimized capture with proper DPI and coordinate handling for high-DPI displays"""
        try:
            lazy_import_pil()
            
            # REDUCED DELAY - much faster capture
            time.sleep(0.05)  # Changed from 0.3 seconds to 0.05 seconds
            
            # Capture full screen - PIL handles DPI automatically
            screenshot = ImageGrab.grab(all_screens=True)
            print(f"Debug: Screenshot size: {screenshot.size}")
            
            # Convert canvas coordinates to screen coordinates
            screen_points = []
            for canvas_x, canvas_y in self.polygon_points:
                # Canvas coordinates are in logical pixels
                logical_x = canvas_x + self.virtual_x
                logical_y = canvas_y + self.virtual_y
                
                # For high-DPI displays, PIL screenshot coordinates might need scaling
                # PIL typically captures at physical resolution
                if screenshot.size[0] > self.virtual_width:
                    # High-DPI mode - scale coordinates
                    actual_scale_x = screenshot.size[0] / self.virtual_width
                    actual_scale_y = screenshot.size[1] / self.virtual_height
                    
                    img_x = int(logical_x * actual_scale_x)
                    img_y = int(logical_y * actual_scale_y)
                else:
                    # Normal DPI mode
                    img_x = logical_x
                    img_y = logical_y
                
                # Handle negative virtual coordinates (multi-monitor)
                if self.virtual_x < 0:
                    img_x = img_x - self.virtual_x
                if self.virtual_y < 0:
                    img_y = img_y - self.virtual_y
                    
                screen_points.append((img_x, img_y))
                print(f"Debug: Canvas ({canvas_x}, {canvas_y}) -> Logical ({logical_x}, {logical_y}) -> Image ({img_x}, {img_y})")

            # Create mask with same size as screenshot
            mask = Image.new('L', screenshot.size, 0)
            
            # Ensure we have valid polygon points
            if len(screen_points) >= 3:
                # Clamp coordinates to image bounds
                clamped_points = []
                for x, y in screen_points:
                    x = max(0, min(x, screenshot.size[0] - 1))
                    y = max(0, min(y, screenshot.size[1] - 1))
                    clamped_points.append((x, y))
                
                ImageDraw.Draw(mask).polygon(clamped_points, fill=255)
                print(f"Debug: Drawing polygon with {len(clamped_points)} points")
            else:
                raise ValueError("Need at least 3 points for polygon")

            # Create result image
            result = Image.new('RGBA', screenshot.size, (0, 0, 0, 0))
            result.paste(screenshot.convert('RGBA'), mask=mask)
            
            # Get bounding box and crop
            bbox = mask.getbbox()
            if bbox and bbox[2] > bbox[0] and bbox[3] > bbox[1]:
                final_image = result.crop(bbox)
                print(f"Debug: Final image size: {final_image.size}")
                self._copy_to_clipboard_fast(final_image)
            else:
                raise ValueError("Invalid polygon area - no content captured")

        except Exception as e:
            print(f"Debug: Capture error: {e}")
            messagebox.showerror("Capture Failed", f"Error: {e}")
        finally:
            self._cleanup_capture()

    def _copy_to_clipboard_fast(self, image):
        """Fast clipboard copy with debug save"""
        # Always save a copy for debugging
        debug_filename = f"debug_screenshot_{int(time.time())}.png"
        image.save(debug_filename)
        print(f"Debug: Saved screenshot as {debug_filename}")
        
        win32clipboard, io = lazy_import_win32()
        
        if not win32clipboard:
            filename = f"polygon_screenshot_{int(time.time())}.png"
            image.save(filename)
            return

        try:
            import io as io_module
            output = io_module.BytesIO()
            image.save(output, "PNG")
            png_data = output.getvalue()

            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            
            try:
                png_format = win32clipboard.RegisterClipboardFormat("PNG")
                win32clipboard.SetClipboardData(png_format, png_data)
                print("Debug: Successfully copied to clipboard as PNG")
            except:
                output = io_module.BytesIO()
                rgb_img = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'RGBA':
                    rgb_img.paste(image, mask=image.split()[-1])
                else:
                    rgb_img = image.convert('RGB')
                
                rgb_img.save(output, "BMP")
                data = output.getvalue()[14:]
                win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
                print("Debug: Successfully copied to clipboard as BMP")

        except Exception as e:
            print(f"Debug: Clipboard error: {e}")
        finally:
            try:
                win32clipboard.CloseClipboard()
            except:
                pass

    def clear_all(self, event):
        """Clear everything"""
        self.canvas.delete("all")
        self.polygon_points = []
        self.is_closed = False

    def cancel_capture(self, event=None):
        """Cancel current capture"""
        self._cleanup_capture()

    def _cleanup_capture(self):
        """Clean up after capture"""
        if self.root:
            self.root.quit()
            self.root = None
        self.active = False

    def stop(self):
        """Stop the service"""
        self.running = False

class PolygonScreenshotService:
    """Main service class that manages background operation"""
    
    def __init__(self):
        self.screenshot_tool = FastPolygonScreenshotTool()
        self.tray_icon = SystemTrayIcon(self.screenshot_tool)
        self.hotkey_thread = None
        
    def setup_hotkeys(self):
        """Setup global hotkeys"""
        keyboard = lazy_import_keyboard()
        
        # Register hotkey: Ctrl+Alt+2
        keyboard.add_hotkey('ctrl+alt+2', self.trigger_capture)
        
        # Keep hotkeys active
        def hotkey_loop():
            while self.screenshot_tool.running:
                time.sleep(0.1)
        
        self.hotkey_thread = threading.Thread(target=hotkey_loop, daemon=True)
        self.hotkey_thread.start()

    def trigger_capture(self):
        """Trigger screenshot capture from hotkey"""
        if not self.screenshot_tool.active:
            threading.Thread(target=self.screenshot_tool.start_capture, daemon=True).start()

    def run(self):
        """Run the service"""
        try:
            # Setup hotkeys
            self.setup_hotkeys()
            
            # Start minimized to tray - no info messages
            # Start the main loop
            self.tray_icon.root.mainloop()
            
        except Exception as e:
            messagebox.showerror("Service Error", f"Failed to start service: {e}")

def main():
    """Main entry point"""
    try:
        service = PolygonScreenshotService()
        service.run()
    except Exception as e:
        messagebox.showerror("Error", f"Critical error: {e}")

if __name__ == "__main__":
    main()