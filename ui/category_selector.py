"""
Category selector dialog
"""
import tkinter as tk
from tkinter import ttk


class CategorySelector:
    """Category selection popup dialog"""
    
    def __init__(self, parent, event, key, current_category, categories, callback):
        self.parent = parent
        self.key = key
        self.current_category = current_category
        self.categories = categories
        self.callback = callback
        
        # Create popup window
        self.popup = tk.Toplevel(parent)
        self.popup.title(f"Change Category for '{key}'")
        self.popup.geometry("350x500")
        self.popup.resizable(False, True)
        
        # Position popup near mouse cursor
        self.popup.geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
        
        # Make popup modal
        self.popup.transient(parent)
        self.popup.grab_set()
        
        self._create_widgets()
        self._setup_bindings()
        self._set_initial_focus()

    def _create_widgets(self):
        """Create all widgets for the dialog"""
        # Header label
        header_label = tk.Label(
            self.popup, 
            text=f"Select category for: {self.key}",
            font=('Arial', 10, 'bold'),
            pady=10
        )
        header_label.pack()
        
        # Current category label
        current_label = tk.Label(
            self.popup,
            text=f"Current: {self.current_category}",
            font=('Arial', 9),
            fg='blue'
        )
        current_label.pack()
        
        # New category input frame
        new_category_frame = tk.Frame(self.popup)
        new_category_frame.pack(fill=tk.X, padx=10, pady=(5, 10))
        
        # New category label and entry
        new_category_label = tk.Label(
            new_category_frame,
            text="Add new category:",
            font=('Arial', 9)
        )
        new_category_label.pack(anchor=tk.W)
        
        self.new_category_entry = tk.Entry(
            new_category_frame,
            font=('Arial', 9),
            width=35
        )
        self.new_category_entry.pack(fill=tk.X, pady=(2, 0))
        
        # Create listbox with scrollbar
        list_label = tk.Label(
            self.popup,
            text="Or select existing category:",
            font=('Arial', 9)
        )
        list_label.pack(anchor=tk.W, padx=10)
        
        list_frame = tk.Frame(self.popup)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))
        
        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            font=('Arial', 9),
            selectmode=tk.SINGLE
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.listbox.yview)
        
        # Populate listbox with categories
        for category in self.categories:
            self.listbox.insert(tk.END, category)
    
        # Find and select current category
        self.current_index = -1
        try:
            self.current_index = self.categories.index(self.current_category)
            self.listbox.selection_set(self.current_index)
            self.listbox.see(self.current_index)
            self.listbox.activate(self.current_index)
        except ValueError:
            self.current_index = -1

        # Buttons frame
        button_frame = tk.Frame(self.popup)
        button_frame.pack(pady=10)
    
        # Change button
        change_btn = tk.Button(
            button_frame,
            text="Change",
            command=self._on_change,
            bg='green',
            fg='white',
            width=12
        )
        change_btn.pack(side=tk.LEFT, padx=5)
    
        # Cancel button
        cancel_btn = tk.Button(
            button_frame,
            text="Cancel",
            command=self._on_cancel,
            width=12
        )
        cancel_btn.pack(side=tk.LEFT, padx=5)

    def _setup_bindings(self):
        """Setup event bindings"""
        # Function to handle selection from listbox
        def on_listbox_select(event):
            if self.listbox.curselection():
                self.new_category_entry.delete(0, tk.END)
    
        # Function to handle typing in entry
        def on_entry_change(event):
            self.listbox.selection_clear(0, tk.END)
    
        # Bind events
        self.listbox.bind('<<ListboxSelect>>', on_listbox_select)
        self.new_category_entry.bind('<KeyRelease>', on_entry_change)
        
        # Keyboard bindings
        def on_double_click(event):
            self._on_change()
    
        def on_key_press(event):
            if event.keysym == 'Return':
                self._on_change()
            elif event.keysym == 'Escape':
                self._on_cancel()
    
        self.listbox.bind('<Double-Button-1>', on_double_click)
        self.new_category_entry.bind('<Return>', lambda e: self._on_change())
        self.popup.bind('<Key>', on_key_press)

    def _set_initial_focus(self):
        """Set initial focus based on whether current category was found"""
        if self.current_index >= 0:
            # Current category found in list - focus on listbox
            self.listbox.focus_set()
        else:
            # Current category not found in list - focus on entry field with current category pre-filled
            self.new_category_entry.insert(0, self.current_category)
            self.new_category_entry.select_range(0, tk.END)
            self.new_category_entry.focus_set()

    def _get_selected_category(self):
        """Get the selected category (either from listbox or entry)"""
        # Check if there's text in the new category entry
        new_category_text = self.new_category_entry.get().strip()
        if new_category_text:
            return new_category_text
    
        # Otherwise get from listbox selection
        selection = self.listbox.curselection()
        if selection:
            return self.categories[selection[0]]
    
        return None

    def _on_change(self):
        """Handle change button click"""
        selected_category = self._get_selected_category()
        if selected_category:
            self.callback(self.key, selected_category)
            self.popup.destroy()
        else:
            # Show error message
            error_label = tk.Label(
                self.popup,
                text="Please select a category or enter a new one",
                fg='red',
                font=('Arial', 8)
            )
            error_label.pack()
            # Remove error after 3 seconds
            self.popup.after(3000, error_label.destroy)

    def _on_cancel(self):
        """Handle cancel button click"""
        self.popup.destroy()