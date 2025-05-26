#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Office Document Translator - GUI Version
========================================
User-friendly executable interface for non-technical users.
Supports Excel, Word, PowerPoint translation between 6 languages.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
import sys
import subprocess
from pathlib import Path
import webbrowser
from typing import List, Optional
import datetime

# Add the current directory to path to import translator
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

try:
    from translator import process_directory, process_file
    TRANSLATOR_AVAILABLE = True
except ImportError as e:
    TRANSLATOR_AVAILABLE = False
    IMPORT_ERROR = str(e)

class OfficeTranslatorGUI:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.setup_variables()
        self.setup_ui()
        self.check_initial_setup()
        
    def setup_window(self):
        """Configure main window"""
        self.root.title("Office Document Translator - Enhanced Edition v2.1")
        self.root.geometry("600x700")
        self.root.resizable(True, True)
        
        # Center window on screen
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (600 // 2)
        y = (self.root.winfo_screenheight() // 2) - (700 // 2)
        self.root.geometry(f"600x700+{x}+{y}")
        
        # Set icon if available
        try:
            # Try to set a default icon
            self.root.iconbitmap("default")
        except:
            pass  # Icon not available, continue without it
            
    def setup_variables(self):
        """Initialize tkinter variables"""
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar(value=os.path.join(current_dir, "output"))
        self.target_language = tk.StringVar(value="ja")
        self.api_status = tk.StringVar(value="❌ Setup Needed")
        self.progress_var = tk.DoubleVar()
        self.status_text = tk.StringVar(value="Ready to translate documents")
        
        # Language options with descriptions
        self.languages = {
            "ja": "🇯🇵 Japanese - Business & Technical Documents",
            "vi": "🇻🇳 Vietnamese - Southeast Asian Market", 
            "en": "🇬🇧 English - Global Business Standard",
            "th": "🇹🇭 Thai - Thailand Market",
            "zh": "🇨🇳 Chinese (Simplified) - China Market",
            "ko": "🇰🇷 Korean - Korea Market"
        }
        
        # Set default input and output paths
        self.input_path.set(os.path.join(current_dir, "input"))
        
        # Ensure directories exist
        os.makedirs(self.input_path.get(), exist_ok=True)
        os.makedirs(self.output_path.get(), exist_ok=True)
        
    def setup_ui(self):
        """Create the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Office Document Translator", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 5))
        
        subtitle_label = ttk.Label(main_frame, text="Enhanced Edition v2.1 - Streamlined", 
                                  font=("Arial", 10))
        subtitle_label.grid(row=1, column=0, columnspan=3, pady=(0, 20))
        
        # Input section
        input_frame = ttk.LabelFrame(main_frame, text="📁 Input Files", padding="10")
        input_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(1, weight=1)
        
        ttk.Label(input_frame, text="Folder:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.input_entry = ttk.Entry(input_frame, textvariable=self.input_path, state="readonly")
        self.input_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(input_frame, text="Browse...", command=self.browse_input).grid(row=0, column=2)
        
        # Drag and drop info
        ttk.Label(input_frame, text="Supported: .xlsx, .xls, .docx, .doc, .pptx, .ppt", 
                 font=("Arial", 9), foreground="gray").grid(row=1, column=0, columnspan=3, pady=(5, 0))
        
        # Output section
        output_frame = ttk.LabelFrame(main_frame, text="💾 Output Location", padding="10")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Label(output_frame, text="Folder:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.output_entry = ttk.Entry(output_frame, textvariable=self.output_path, state="readonly")
        self.output_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(output_frame, text="Browse...", command=self.browse_output).grid(row=0, column=2)
        
        # Language selection
        lang_frame = ttk.LabelFrame(main_frame, text="🌍 Translation Language", padding="10")
        lang_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        lang_frame.columnconfigure(1, weight=1)
        
        ttk.Label(lang_frame, text="Translate to:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.lang_combo = ttk.Combobox(lang_frame, textvariable=self.target_language,
                                      values=list(self.languages.keys()), state="readonly")
        self.lang_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        self.lang_combo.bind('<<ComboboxSelected>>', self.on_language_change)
        
        # Language description
        self.lang_desc = ttk.Label(lang_frame, text=self.languages["ja"], font=("Arial", 9))
        self.lang_desc.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        
        # API Status section
        api_frame = ttk.LabelFrame(main_frame, text="🔑 API Configuration", padding="10")
        api_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        api_frame.columnconfigure(1, weight=1)
        
        ttk.Label(api_frame, text="Status:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.api_label = ttk.Label(api_frame, textvariable=self.api_status)
        self.api_label.grid(row=0, column=1, sticky=tk.W, padx=(0, 10))
        ttk.Button(api_frame, text="Setup API Key", command=self.setup_api_key).grid(row=0, column=2)
        
        # Quick setup link
        setup_link = ttk.Label(api_frame, text="Get FREE Gemini API Key", 
                              foreground="blue", cursor="hand2", font=("Arial", 9, "underline"))
        setup_link.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        setup_link.bind("<Button-1>", lambda e: webbrowser.open("https://aistudio.google.com/app/apikey"))
        
        # Action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=20)
        
        self.translate_btn = ttk.Button(button_frame, text="🚀 Start Translation", 
                                       command=self.start_translation, style="Accent.TButton")
        self.translate_btn.grid(row=0, column=0, padx=(0, 10))
        
        ttk.Button(button_frame, text="📁 Open Output", command=self.open_output).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(button_frame, text="📋 View Files", command=self.view_files).grid(row=0, column=2)
        
        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="📊 Translation Progress", padding="10")
        progress_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, style="TProgressbar")
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.status_label = ttk.Label(progress_frame, textvariable=self.status_text)
        self.status_label.grid(row=1, column=0, sticky=tk.W)
        
        # Log section
        log_frame = ttk.LabelFrame(main_frame, text="📝 Activity Log", padding="10")
        log_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(8, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=8, state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Initialize
        self.log_message("✅ Office Document Translator initialized successfully")
        self.log_message("📁 Input folder: " + self.input_path.get())
        self.log_message("💾 Output folder: " + self.output_path.get())
        
    def browse_input(self):
        """Browse for input folder"""
        folder = filedialog.askdirectory(title="Select Input Folder", 
                                        initialdir=self.input_path.get())
        if folder:
            self.input_path.set(folder)
            self.log_message(f"📁 Input folder changed to: {folder}")
            
    def browse_output(self):
        """Browse for output folder"""
        folder = filedialog.askdirectory(title="Select Output Folder", 
                                        initialdir=self.output_path.get())
        if folder:
            self.output_path.set(folder)
            os.makedirs(folder, exist_ok=True)
            self.log_message(f"💾 Output folder changed to: {folder}")
            
    def on_language_change(self, event=None):
        """Handle language selection change"""
        lang_code = self.target_language.get()
        self.lang_desc.config(text=self.languages.get(lang_code, ""))
        lang_name = self.languages.get(lang_code, "").split(" - ")[0]
        self.log_message(f"🌍 Target language set to: {lang_name}")
        
    def setup_api_key(self):
        """Show API key setup dialog"""
        dialog = APIKeyDialog(self.root, self.check_api_status)
        
    def check_api_status(self):
        """Check if API key is configured"""
        env_file = os.path.join(current_dir, ".env")
        if os.path.exists(env_file):
            try:
                with open(env_file, 'r') as f:
                    content = f.read()
                    if "GEMINI_API_KEY=" in content and len(content.split("GEMINI_API_KEY=")[1].strip()) > 10:
                        self.api_status.set("✅ API Key Configured")
                        self.log_message("🔑 API key detected and configured")
                        return True
            except Exception as e:
                pass
                
        self.api_status.set("❌ Setup Needed")
        return False
        
    def check_initial_setup(self):
        """Check initial setup status"""
        if not TRANSLATOR_AVAILABLE:
            self.log_message("❌ Translation engine not available")
            self.log_message(f"Error: {IMPORT_ERROR}")
            self.translate_btn.config(state="disabled")
            messagebox.showerror("Setup Error", 
                               f"Translation engine not available.\n\nError: {IMPORT_ERROR}\n\n"
                               "Please ensure all dependencies are installed.")
        else:
            self.log_message("✅ Translation engine loaded successfully")
            
        self.check_api_status()
        
    def view_files(self):
        """Show files in input directory"""
        input_dir = self.input_path.get()
        if not os.path.exists(input_dir):
            messagebox.showwarning("Directory Not Found", f"Input directory does not exist:\n{input_dir}")
            return
            
        # Count files by type
        excel_files = list(Path(input_dir).glob("*.xlsx")) + list(Path(input_dir).glob("*.xls"))
        word_files = list(Path(input_dir).glob("*.docx")) + list(Path(input_dir).glob("*.doc"))
        ppt_files = list(Path(input_dir).glob("*.pptx")) + list(Path(input_dir).glob("*.ppt"))
        
        total_files = len(excel_files) + len(word_files) + len(ppt_files)
        
        if total_files == 0:
            messagebox.showinfo("No Files Found", 
                              f"No supported Office documents found in:\n{input_dir}\n\n"
                              "Supported formats: .xlsx, .xls, .docx, .doc, .pptx, .ppt")
        else:
            file_info = f"Found {total_files} document(s):\n\n"
            file_info += f"📊 Excel: {len(excel_files)} files\n"
            file_info += f"📄 Word: {len(word_files)} files\n"
            file_info += f"📽️ PowerPoint: {len(ppt_files)} files\n\n"
            file_info += f"Location: {input_dir}"
            
            messagebox.showinfo("Files Ready for Translation", file_info)
            
    def open_output(self):
        """Open output folder in file explorer"""
        output_dir = self.output_path.get()
        if os.path.exists(output_dir):
            try:
                if sys.platform == "win32":
                    os.startfile(output_dir)
                elif sys.platform == "darwin":  # macOS
                    subprocess.run(["open", output_dir])
                else:  # Linux
                    subprocess.run(["xdg-open", output_dir])
                self.log_message(f"📁 Opened output folder: {output_dir}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not open folder:\n{str(e)}")
        else:
            messagebox.showwarning("Folder Not Found", f"Output folder does not exist:\n{output_dir}")
            
    def log_message(self, message: str):
        """Add message to log"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}\n"
        
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        
        # Update status
        self.status_text.set(message)
        self.root.update_idletasks()
        
    def start_translation(self):
        """Start translation process in background thread"""
        if not TRANSLATOR_AVAILABLE:
            messagebox.showerror("Error", "Translation engine not available. Please check setup.")
            return
            
        if not self.check_api_status():
            messagebox.showwarning("API Key Required", "Please configure your Gemini API key before translating.")
            self.setup_api_key()
            return
            
        input_dir = self.input_path.get()
        if not os.path.exists(input_dir):
            messagebox.showerror("Error", f"Input directory does not exist:\n{input_dir}")
            return
            
        # Disable translation button
        self.translate_btn.config(state="disabled", text="🔄 Translating...")
        self.progress_var.set(0)
        
        # Start translation in background
        thread = threading.Thread(target=self.run_translation, daemon=True)
        thread.start()
        
    def run_translation(self):
        """Run translation process"""
        try:
            input_dir = self.input_path.get()
            output_dir = self.output_path.get()
            target_lang = self.target_language.get()
            
            self.log_message("🚀 Starting translation process...")
            self.log_message(f"📁 Input: {input_dir}")
            self.log_message(f"💾 Output: {output_dir}")
            self.log_message(f"🌍 Target: {self.languages[target_lang].split(' - ')[0]}")
            
            # Ensure output directory exists
            os.makedirs(output_dir, exist_ok=True)
            
            # Use the translator module
            self.progress_var.set(10)
            successful, failed = process_directory(input_dir, target_lang)
            
            self.progress_var.set(100)
            
            # Show results
            if successful:
                result_msg = f"✅ Translation completed successfully!\n\n"
                result_msg += f"📊 Processed: {len(successful)} file(s)\n"
                if failed:
                    result_msg += f"❌ Failed: {len(failed)} file(s)\n"
                result_msg += f"📁 Output location: {output_dir}"
                
                self.log_message(f"✅ Translation completed! {len(successful)} files processed")
                if failed:
                    self.log_message(f"⚠️ {len(failed)} files failed to process")
                    
                messagebox.showinfo("Translation Complete", result_msg)
            else:
                error_msg = "❌ No files were successfully translated.\n\n"
                if failed:
                    error_msg += f"Failed files: {len(failed)}\n"
                error_msg += "Please check the log for details."
                
                self.log_message("❌ Translation failed - no files processed")
                messagebox.showerror("Translation Failed", error_msg)
                
        except Exception as e:
            self.log_message(f"❌ Translation error: {str(e)}")
            messagebox.showerror("Translation Error", f"An error occurred during translation:\n\n{str(e)}")
            
        finally:
            # Re-enable translation button
            self.root.after(0, self.reset_translation_button)
            
    def reset_translation_button(self):
        """Reset translation button state"""
        self.translate_btn.config(state="normal", text="🚀 Start Translation")
        self.progress_var.set(0)
        self.status_text.set("Ready to translate documents")


class APIKeyDialog:
    def __init__(self, parent, callback=None):
        self.parent = parent
        self.callback = callback
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Gemini API Key Setup")
        self.dialog.geometry("500x300")
        self.dialog.resizable(False, False)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center dialog
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (500 // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (300 // 2)
        self.dialog.geometry(f"500x300+{x}+{y}")
        
        self.setup_ui()
        
    def setup_ui(self):
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="🔑 Gemini API Key Setup", 
                               font=("Arial", 14, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Instructions
        instructions = """To use the Office Document Translator, you need a FREE Gemini API key.

1. Click the link below to get your API key
2. Copy your API key 
3. Paste it in the field below
4. Click Save

Your API key will be stored securely in a local .env file."""
        
        inst_label = ttk.Label(main_frame, text=instructions, justify=tk.LEFT)
        inst_label.pack(pady=(0, 15))
        
        # Link
        link_label = ttk.Label(main_frame, text="🌐 Get FREE Gemini API Key", 
                              foreground="blue", cursor="hand2", 
                              font=("Arial", 10, "underline"))
        link_label.pack(pady=(0, 20))
        link_label.bind("<Button-1>", lambda e: webbrowser.open("https://aistudio.google.com/app/apikey"))
        
        # API Key input
        key_frame = ttk.Frame(main_frame)
        key_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(key_frame, text="API Key:").pack(anchor=tk.W)
        self.api_key_var = tk.StringVar()
        
        # Load existing key if available
        self.load_existing_key()
        
        self.api_entry = ttk.Entry(key_frame, textvariable=self.api_key_var, show="*", width=60)
        self.api_entry.pack(fill=tk.X, pady=(5, 0))
        
        # Show/hide button
        show_frame = ttk.Frame(main_frame)
        show_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.show_key = tk.BooleanVar()
        show_check = ttk.Checkbutton(show_frame, text="Show API key", variable=self.show_key,
                                    command=self.toggle_key_visibility)
        show_check.pack(anchor=tk.W)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="Save", command=self.save_key).pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Button(button_frame, text="Cancel", command=self.dialog.destroy).pack(side=tk.RIGHT)
        
        # Focus on entry
        self.api_entry.focus()
        
    def load_existing_key(self):
        """Load existing API key if available"""
        env_file = os.path.join(current_dir, ".env")
        if os.path.exists(env_file):
            try:
                with open(env_file, 'r') as f:
                    content = f.read()
                    if "GEMINI_API_KEY=" in content:
                        key = content.split("GEMINI_API_KEY=")[1].strip()
                        self.api_key_var.set(key)
            except Exception:
                pass
                
    def toggle_key_visibility(self):
        """Toggle API key visibility"""
        if self.show_key.get():
            self.api_entry.config(show="")
        else:
            self.api_entry.config(show="*")
            
    def save_key(self):
        """Save API key to .env file"""
        api_key = self.api_key_var.get().strip()
        
        if not api_key:
            messagebox.showerror("Error", "Please enter your API key.")
            return
            
        if len(api_key) < 10:
            messagebox.showerror("Error", "API key appears to be too short. Please check and try again.")
            return
            
        try:
            env_file = os.path.join(current_dir, ".env")
            with open(env_file, 'w') as f:
                f.write(f"GEMINI_API_KEY={api_key}\n")
                
            messagebox.showinfo("Success", "API key saved successfully!")
            
            if self.callback:
                self.callback()
                
            self.dialog.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save API key:\n{str(e)}")


def main():
    """Main application entry point"""
    # Check if this is running as exe or script
    if getattr(sys, 'frozen', False):
        # Running as exe
        application_path = os.path.dirname(sys.executable)
    else:
        # Running as script
        application_path = os.path.dirname(os.path.abspath(__file__))
    
    # Change to application directory
    os.chdir(application_path)
    
    # Create main window
    root = tk.Tk()
    
    # Set up style
    style = ttk.Style()
    try:
        # Try to use a modern theme
        style.theme_use('vista')  # Windows
    except:
        try:
            style.theme_use('clam')  # Cross-platform
        except:
            pass  # Use default theme
    
    # Create and run application
    app = OfficeTranslatorGUI(root)
    
    # Handle window close
    def on_closing():
        root.quit()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main() 