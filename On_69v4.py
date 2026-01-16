import tkinter as tk
from tkinter import ttk
import pyautogui
import win32gui
import win32api
import win32con

class ProfessionalKeyboard:
    def __init__(self, root):
        self.root = root
        self.root.title("Professional On-Screen Keyboard")
        self.root.geometry("900x400")
        self.root.attributes('-topmost', True)  # Keep the keyboard on top
        self.root.configure(bg='#1e1e1e')
        self.root.resizable(False, False)

        # Style configurationhh
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Color scheme
        self.bg_color = '#1e1e1e'
        self.button_bg = '#2d2d2d'
        self.button_fg = '#ffffff'
        self.button_en = '#008000'
        self.special_button_bg = '#007acc'
        self.danger_button_bg = '#d9534f'  # Red color for backspace
        self.hover_color = '#696969'
        self.display_bg = '#252526'
        self.display_fg = '#00ff00'

        # Configure styles
        self.style.configure('TButton',
                             font=('Arial', 14, 'bold'),
                             borderwidth=2,
                             relief='flat',
                             background=self.button_bg,
                             foreground=self.button_fg)

        self.style.map('TButton',
                       background=[('active', self.hover_color)],
                       foreground=[('active', self.button_fg)])

        self.style.configure('Special.TButton',
                             background=self.button_bg,
                             foreground='white')

        self.style.configure('Entry.TButton',
                             background=self.special_button_bg,
                             foreground='white')

        self.style.configure('Enter.TButton',
                             background=self.button_en,
                             foreground='white')

        self.style.configure('Danger.TButton',  # Style for backspace
                             background=self.danger_button_bg,
                             foreground='white')

        self.style.configure('Display.TEntry',
                             font=('Arial', 24),
                             fieldbackground=self.display_bg,
                             foreground=self.display_fg,
                             padding=10,
                             borderwidth=3,
                             relief='sunken')

        # Variables to track caret position
        self.last_caret_pos = pyautogui.position()  # Initialize with the current mouse position
        self.previous_window = win32gui.GetForegroundWindow()  # Store the previously active window
        self.caps_lock_on = False

        # Create UI components
        self.create_display()
        self.create_keyboard()

        # Bind mouse events to track clicks and update caret position
        self.root.bind("<Button-1>", self.start_drag)
        self.root.bind("<B1-Motion>", self.drag_window)

        # Periodically track mouse clicks outside the keyboard
        self.root.after(100, self.track_mouse_clicks)

    def track_mouse_clicks(self):
        """
        Dynamically track the most recent mouse click position
        and store it as the current caret position.
        """
        try:
            hwnd = win32gui.GetForegroundWindow()
            if hwnd != self.get_hwnd():  # Ensure the click is outside the keyboard
                self.last_caret_pos = pyautogui.position()  # Update the caret position
                self.previous_window = hwnd  # Update the active window
        except Exception as e:
            print(f"Error tracking mouse clicks: {e}")
        finally:
            # Continue tracking
            self.root.after(100, self.track_mouse_clicks)

    def get_hwnd(self):
        """Get the window handle for the on-screen keyboard."""
        return win32gui.FindWindow(None, self.root.title())

    def create_display(self):
        """Create the display area for the keyboard."""
        display_frame = ttk.Frame(self.root)
        display_frame.pack(fill="x", pady=(20, 10), padx=20)

        self.display_var = tk.StringVar()

        self.display = ttk.Entry(
            display_frame,
            textvariable=self.display_var,
            style='Display.TEntry',
            justify="center",
            state="readonly"
        )
        self.display.pack(fill="x")

    def create_keyboard(self):
        """Create the on-screen keyboard buttons."""
        keyboard_frame = ttk.Frame(self.root)
        keyboard_frame.pack(expand=True, fill="both", padx=20, pady=10)

        # Keyboard rows
        keys = [
            ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0'],
            ['Q', 'W', 'E', 'R', 'T', 'Y', 'U', 'I', 'O', 'P'],
            ['A', 'S', 'D', 'F', 'G', 'H', 'J', 'K', 'L'],
            ['Z', 'X', 'C', 'V', 'B', 'N', 'M', ',', '.', '?']
        ]

        # Create rows for alphanumeric keys
        for row_idx, row in enumerate(keys):
            row_frame = ttk.Frame(keyboard_frame)
            row_frame.pack(expand=True, fill="both", pady=2)

            for key in row:
                btn = ttk.Button(
                    row_frame,
                    text=key,
                    style='TButton',
                    width=6,
                    command=lambda k=key: self.handle_key(k)
                )
                btn.pack(side="left", expand=True, fill="both", padx=1)

        # Create a row for special keys
        special_row_frame = ttk.Frame(keyboard_frame)
        special_row_frame.pack(expand=True, fill="both", pady=2)

        # Add the Enter button
        enter_button = ttk.Button(
            special_row_frame,
            text='Enter',
            style='Enter.TButton',
            width=15,
            command=lambda: self.handle_key('Enter')
        )
        enter_button.pack(side="left", expand=True, fill="both", padx=1)

        # Add the Caps Lock button
        caps_button = ttk.Button(
            special_row_frame,
            text='Caps',
            style='Special.TButton',
            width=10,
            command=lambda: self.handle_key('Caps')
        )
        caps_button.pack(side="left", expand=True, fill="both", padx=1)

        # Add the Space button
        space_button = ttk.Button(
            special_row_frame,
            text='Space',
            style='Special.TButton',
            width=15,
            command=lambda: self.handle_key('Space')
        )
        space_button.pack(side="left", expand=True, fill="both", padx=1)

        # Add the Clear All button
        clear_button = ttk.Button(
            special_row_frame,
            text='Clear All',
            style='Special.TButton',
            width=10,
            command=lambda: self.handle_key('Clear')
        )
        clear_button.pack(side="left", expand=True, fill="both", padx=1)

        # Add the Backspace button
        backspace_button = ttk.Button(
            special_row_frame,
            text='⌫',
            style='Danger.TButton',
            width=20,
            command=lambda: self.handle_key('⌫')
        )
        backspace_button.pack(side="left", expand=True, fill="both", padx=1)

    def handle_key(self, key):
        """Handle keypresses and simulate typing at the caret position."""
        current = self.display_var.get()

        # Restore focus to the last clicked window
        win32gui.SetForegroundWindow(self.previous_window)

        if key == '⌫':
            self.display_var.set(current[:-1])  # Remove last character
            pyautogui.press('backspace')  # Simulate backspace
        elif key == 'Enter':
            pyautogui.press('enter')  # Simulate Enter key
            self.display_var.set('')  # Clear display
        elif key == 'Caps':
            self.caps_lock_on = not self.caps_lock_on  # Toggle Caps Lock
        elif key == 'Space':
            self.display_var.set(current + ' ')  # Add space to display
            pyautogui.write(' ')  # Simulate space
        elif key == 'Clear':
            self.display_var.set('')  # Clear display
            pyautogui.hotkey('ctrl', 'a')  # Select all text
            pyautogui.press('delete')  # Delete the selected text
        else:
            char = key.upper() if self.caps_lock_on else key.lower()
            self.display_var.set(current + char)  # Add character to display
            pyautogui.write(char)  # Simulate typing

    def start_drag(self, event):
        """Initiates dragging."""
        self.is_dragging = True
        self.drag_start_x = event.x
        self.drag_start_y = event.y

    def drag_window(self, event):
        """Moves the window while dragging."""
        if self.is_dragging:
            x = self.root.winfo_x() - self.drag_start_x + event.x
            y = self.root.winfo_y() - self.drag_start_y + event.y
            self.root.geometry(f"+{x}+{y}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ProfessionalKeyboard(root)
    root.mainloop()
