import os
import sys
import subprocess
import json
import threading
import time
import re
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Tuple, Set

# GUI libraries
import customtkinter as ctk
from tkinter import filedialog, StringVar, BooleanVar, messagebox

# Set appearance mode and default theme
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Simplified language mapping (ISO 639-2 to display name)
LANGUAGE_MAPPING = {
    "eng": "English",
    "heb": "Hebrew",
    "und": "Undefined"
}

# Supported output formats
SUBTITLE_FORMATS = {
    "SRT": {
        "extension": "srt",
        "ffmpeg_format": "srt",
        "description": "SubRip Text"
    },
    "VobSub": {
        "extension": "idx",  # .idx and .sub will be created
        "ffmpeg_format": "dvdsub",
        "description": "DVD Subtitles"
    },
    "ASS": {
        "extension": "ass",
        "ffmpeg_format": "ass",
        "description": "Advanced SubStation Alpha"
    },
    "WebVTT": {
        "extension": "vtt",
        "ffmpeg_format": "webvtt",
        "description": "Web Video Text Tracks"
    }
}

class SubtitleExtractor:
    """Core functionality for extracting subtitles from video files"""
    
    def __init__(self, log_callback=None):
        self.log_callback = log_callback
        self.cancel_flag = False
        self.total_streams = 0
        self.processed_streams = 0
        self.successful_extractions = 0
        self.failed_extractions = 0
        self.skipped_extractions = 0
    
    def log(self, message):
        """Log messages to the callback if provided"""
        if self.log_callback:
            self.log_callback(message)
        else:
            print(message)
    
    def check_dependencies(self) -> bool:
        """Check if ffmpeg and ffprobe are available in the system"""
        try:
            subprocess.run(
                ["ffmpeg", "-version"], 
                stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE,
                check=True
            )
            subprocess.run(
                ["ffprobe", "-version"], 
                stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE,
                check=True
            )
            return True
        except (subprocess.SubprocessError, FileNotFoundError):
            self.log("ERROR: ffmpeg or ffprobe not found. Please install them and make sure they're in your PATH.")
            return False
    
    def get_video_files(self, directory: str) -> List[str]:
        """Get all video files from directory"""
        video_extensions = ('.mkv', '.mp4', '.avi', '.mov', '.wmv', '.flv', '.webm')
        video_files = []
        
        for root, _, files in os.walk(directory):
            for file in files:
                if file.lower().endswith(video_extensions):
                    video_files.append(os.path.join(root, file))
        
        return video_files
    
    def get_subtitle_streams(self, video_file: str) -> List[Dict]:
        """Get all subtitle streams from video file"""
        try:
            # Run ffprobe to get stream information in JSON format
            result = subprocess.run(
                [
                    "ffprobe", 
                    "-v", "quiet", 
                    "-print_format", "json", 
                    "-show_streams", 
                    "-select_streams", "s", 
                    video_file
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                errors='replace',  # Handle Unicode decode errors
                check=False
            )
            
            # Parse JSON output
            streams_data = json.loads(result.stdout)
            subtitle_streams = []
            
            # Extract relevant subtitle stream information
            if 'streams' in streams_data:
                for stream in streams_data['streams']:
                    stream_info = {
                        'index': stream.get('index'),
                        'codec_name': stream.get('codec_name', 'unknown'),
                        'language': stream.get('tags', {}).get('language', 'und'),
                        'title': stream.get('tags', {}).get('title', '')
                    }
                    subtitle_streams.append(stream_info)
            
            return subtitle_streams
        
        except subprocess.SubprocessError as e:
            self.log(f"ERROR: Failed to get subtitle streams from {os.path.basename(video_file)}: {str(e)}")
            return []
        except json.JSONDecodeError:
            self.log(f"ERROR: Failed to parse ffprobe output for {os.path.basename(video_file)}")
            return []
    
    def extract_subtitle(
        self,
        video_file: str,
        stream_index: int,
        language: str,
        output_format: str,
        output_dir: Optional[str] = None,
        overwrite: bool = False
    ) -> bool:
        """
        Extract a subtitle stream from a video file
        
        Args:
            video_file: Path to the video file
            stream_index: Index of the subtitle stream to extract
            language: Language code of the subtitle
            output_format: Format to extract to (e.g., "SRT", "VobSub")
            output_dir: Directory to save the subtitle file (default: same as video)
            overwrite: Whether to overwrite existing subtitle files
            
        Returns:
            bool: True if extraction was successful
        """
        # Get output directory and base filename
        video_basename = os.path.splitext(os.path.basename(video_file))[0]
        if output_dir is None:
            output_dir = os.path.dirname(video_file)
            
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Get format information
        if output_format not in SUBTITLE_FORMATS:
            self.log(f"ERROR: Unsupported format {output_format}")
            self.failed_extractions += 1
            return False
            
        format_info = SUBTITLE_FORMATS[output_format]
        extension = format_info["extension"]
        ffmpeg_format = format_info["ffmpeg_format"]
        
        # Create output filename
        output_file = os.path.join(output_dir, f"{video_basename}.{language}.{extension}")
        
        # Check if file already exists
        if os.path.exists(output_file) and not overwrite:
            self.log(f"SKIPPED: Subtitle file already exists for {video_basename} [{language}]")
            self.skipped_extractions += 1
            return False
            
        # Build ffmpeg command based on format
        cmd = [
            "ffmpeg",
            "-loglevel", "warning",  # Reduce verbosity
            "-y" if overwrite else "-n",  # Overwrite or not
            "-i", video_file,
            "-map", f"0:{stream_index}",
            "-c:s", ffmpeg_format,
        ]
        
        # Add format-specific parameters
        if output_format == "VobSub":
            # VobSub requires both .idx and .sub files
            cmd.extend([output_file])
        else:
            # Other formats use a single file
            cmd.extend([output_file])
        
        try:
            # Run ffmpeg command
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                errors='replace'  # Handle Unicode decode errors
            )
            
            # Read and parse stderr output for progress information
            stderr_lines = []
            for line in process.stderr:
                stderr_lines.append(line)
                if "time=" in line:
                    # Could extract progress here if needed
                    pass
                    
                # Check for cancel flag
                if self.cancel_flag:
                    process.terminate()
                    self.log(f"CANCELLED: Extraction cancelled for {video_basename} [{language}]")
                    return False
            
            # Wait for process to complete
            return_code = process.wait()
            
            if return_code != 0:
                error_output = ''.join(stderr_lines)
                
                # Check for common errors and provide more specific messages
                if "Unknown encoder" in error_output:
                    self.log(f"ERROR: Format mismatch or unsupported format for {video_basename} [{language}]")
                elif "Output file exists" in error_output:
                    self.log(f"SKIPPED: Subtitle file already exists for {video_basename} [{language}]")
                    self.skipped_extractions += 1
                else:
                    self.log(f"ERROR: Failed to extract subtitle from {video_basename} [{language}]")
                    
                self.failed_extractions += 1
                return False
            
            self.log(f"SUCCESS: Extracted {output_format} subtitle for {video_basename} [{language}]")
            self.successful_extractions += 1
            return True
            
        except subprocess.SubprocessError as e:
            self.log(f"ERROR: Failed to run ffmpeg for {video_basename} [{language}]: {str(e)}")
            self.failed_extractions += 1
            return False
    
    def reset(self):
        """Reset the extractor state"""
        self.cancel_flag = False
        self.total_streams = 0
        self.processed_streams = 0
        self.successful_extractions = 0
        self.failed_extractions = 0
        self.skipped_extractions = 0
    
    def cancel(self):
        """Cancel the extraction process"""
        self.cancel_flag = True
        self.log("INFO: Cancellation requested. Waiting for current task to complete...")


class SubtitleExtractorGUI(ctk.CTk):
    """GUI for the subtitle extractor application"""
    
    def __init__(self):
        super().__init__()
        
        # Initialize the subtitle extractor
        self.extractor = SubtitleExtractor(log_callback=self.log_message)
        
        # Setup the GUI
        self.title("Subtitle Extractor")
        self.geometry("900x700")
        self.minsize(800, 600)
        
        # State variables
        self.selected_directory = StringVar(value="")
        self.selected_subdirs = []
        self.language_var = StringVar(value="eng")
        self.format_var = StringVar(value="SRT")
        self.custom_language_var = StringVar(value="")
        self.overwrite_var = BooleanVar(value=False)
        self.custom_output_dir_var = BooleanVar(value=False)
        self.output_directory = StringVar(value="")
        
        # Create layout frames
        self.create_layout()
        
        # Thread for extraction
        self.extraction_thread = None
        self.start_time = None
        
        # Check dependencies
        if not self.extractor.check_dependencies():
            messagebox.showerror(
                "Dependency Error",
                "ffmpeg or ffprobe not found. Please install them and make sure they're in your PATH."
            )
    
    def create_layout(self):
        """Create the application layout"""
        # Create a main frame that will contain all elements in a vertical layout
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Configure the main frame to use a grid layout
        self.main_frame.grid_columnconfigure(0, weight=1)
        
        # Row counter for adding elements
        row = 0
        
        # Top frame for directory selection (row 0)
        self.create_directory_frame(self.main_frame, row)
        row += 1
        
        # Subdirectory selection frame (row 1)
        self.create_subdir_frame(self.main_frame, row)
        row += 1
        
        # Middle frame for options (row 2)
        self.create_options_frame(self.main_frame, row)
        row += 1
        
        # Bottom frame for progress and log (row 3)
        self.create_progress_frame(self.main_frame, row)
        
    def create_directory_frame(self, parent, row):
        """Create the directory selection frame"""
        dir_frame = ctk.CTkFrame(parent)
        dir_frame.grid(row=row, column=0, padx=0, pady=(0, 5), sticky="ew")
        dir_frame.grid_columnconfigure(1, weight=1)
        
        # Directory selection
        ctk.CTkLabel(dir_frame, text="Source Directory:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        dir_entry = ctk.CTkEntry(dir_frame, textvariable=self.selected_directory, width=400)
        dir_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        browse_btn = ctk.CTkButton(dir_frame, text="Browse", command=self.browse_directory)
        browse_btn.grid(row=0, column=2, padx=5, pady=5)
        
    def create_subdir_frame(self, parent, row):
        """Create subdirectory selection frame"""
        # Subdirectory selection frame (initially hidden, will be shown after directory selection)
        self.subdir_frame = ctk.CTkFrame(parent)
        self.subdir_frame.grid(row=row, column=0, padx=0, pady=(0, 5), sticky="nsew")
        self.subdir_frame.grid_columnconfigure(0, weight=1)
        self.subdir_frame.grid_rowconfigure(1, weight=1)
        
        # Label and buttons for subdirectory selection
        subdir_header = ctk.CTkFrame(self.subdir_frame)
        subdir_header.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        
        ctk.CTkLabel(subdir_header, text="Select Subdirectories:").pack(side="left", padx=5, pady=5)
        
        # Action buttons in the header (right aligned)
        buttons_frame = ctk.CTkFrame(subdir_header, fg_color="transparent")
        buttons_frame.pack(side="right", fill="y")
        
        # Apply button for subdirectory selection
        self.apply_btn = ctk.CTkButton(
            buttons_frame, 
            text="Apply Selection", 
            command=self.apply_selection,
            fg_color="#28a745",
            hover_color="#218838",
            width=120
        )
        self.apply_btn.pack(side="left", padx=5, pady=5)
        
        # Select/Deselect buttons
        ctk.CTkButton(
            buttons_frame, 
            text="Deselect All", 
            command=self.deselect_all_subdirs,
            width=100
        ).pack(side="left", padx=5, pady=5)
        
        ctk.CTkButton(
            buttons_frame, 
            text="Select All", 
            command=self.select_all_subdirs,
            width=100
        ).pack(side="left", padx=5, pady=5)
        
        # Scrollable frame for subdirectories (fixed height to prevent overlapping)
        self.subdir_scrollable = ctk.CTkScrollableFrame(self.subdir_frame, height=150)
        self.subdir_scrollable.grid(row=1, column=0, padx=5, pady=5, sticky="ew")
        
        # Hide the subdirectory frame initially
        self.subdir_frame.grid_remove()
        
    def create_options_frame(self, parent, row):
        """Create the options frame"""
        options_frame = ctk.CTkFrame(parent)
        options_frame.grid(row=row, column=0, padx=0, pady=(0, 5), sticky="ew")
        options_frame.grid_columnconfigure(1, weight=1)
        options_frame.grid_columnconfigure(3, weight=1)
        
        # Language selection
        ctk.CTkLabel(options_frame, text="Language:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        language_menu = ctk.CTkOptionMenu(
            options_frame,
            variable=self.language_var,
            values=list(LANGUAGE_MAPPING.keys()),
            dynamic_resizing=False,
            command=self.on_language_change,
            width=150
        )
        language_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # Custom language entry
        ctk.CTkLabel(options_frame, text="Custom Language:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        custom_language_entry = ctk.CTkEntry(options_frame, textvariable=self.custom_language_var, width=150)
        custom_language_entry.grid(row=0, column=3, padx=5, pady=5, sticky="w")
        
        # Format selection
        ctk.CTkLabel(options_frame, text="Output Format:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        format_menu = ctk.CTkOptionMenu(
            options_frame,
            variable=self.format_var,
            values=list(SUBTITLE_FORMATS.keys()),
            dynamic_resizing=False,
            width=150
        )
        format_menu.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        # Overwrite option
        overwrite_check = ctk.CTkCheckBox(
            options_frame,
            text="Overwrite existing files",
            variable=self.overwrite_var
        )
        overwrite_check.grid(row=1, column=2, padx=5, pady=5, sticky="w")
        
        # Custom output directory
        custom_dir_check = ctk.CTkCheckBox(
            options_frame,
            text="Custom output directory",
            variable=self.custom_output_dir_var,
            command=self.toggle_output_dir
        )
        custom_dir_check.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        
        # Output directory selection (initially disabled)
        self.output_dir_frame = ctk.CTkFrame(options_frame)
        self.output_dir_frame.grid(row=2, column=1, columnspan=3, padx=5, pady=5, sticky="ew")
        self.output_dir_frame.grid_columnconfigure(0, weight=1)
        
        output_dir_entry = ctk.CTkEntry(self.output_dir_frame, textvariable=self.output_directory, width=400, state="disabled")
        output_dir_entry.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        self.output_dir_browse_btn = ctk.CTkButton(
            self.output_dir_frame,
            text="Browse",
            command=self.browse_output_dir,
            state="disabled"
        )
        self.output_dir_browse_btn.pack(side="right", padx=5, pady=5)
        
        # Action buttons
        action_frame = ctk.CTkFrame(options_frame)
        action_frame.grid(row=3, column=0, columnspan=4, padx=5, pady=5, sticky="ew")
        
        self.extract_btn = ctk.CTkButton(
            action_frame,
            text="Extract Subtitles",
            command=self.start_extraction,
            width=150
        )
        self.extract_btn.pack(side="left", padx=5, pady=5)
        
        self.cancel_btn = ctk.CTkButton(
            action_frame,
            text="Cancel",
            command=self.cancel_extraction,
            state="disabled",
            fg_color="#FF5555",
            hover_color="#FF0000",
            width=100
        )
        self.cancel_btn.pack(side="left", padx=5, pady=5)
        
    def create_progress_frame(self, parent, row):
        """Create the progress tracking frame"""
        progress_frame = ctk.CTkFrame(parent)
        progress_frame.grid(row=row, column=0, padx=0, pady=(0, 5), sticky="nsew")
        progress_frame.grid_columnconfigure(0, weight=1)
        progress_frame.grid_rowconfigure(1, weight=1)
        
        # Progress indicators
        progress_indicators = ctk.CTkFrame(progress_frame)
        progress_indicators.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        progress_indicators.grid_columnconfigure(1, weight=1)
        
        # Progress bar label
        self.progress_label = ctk.CTkLabel(progress_indicators, text="Ready")
        self.progress_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(progress_indicators)
        self.progress_bar.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.progress_bar.set(0)
        
        # Time information
        self.time_label = ctk.CTkLabel(progress_indicators, text="Elapsed: 00:00 | Remaining: --:--")
        self.time_label.grid(row=0, column=2, padx=5, pady=5, sticky="e")
        
        # Log text area
        log_frame = ctk.CTkFrame(progress_frame)
        log_frame.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)
        
        # Set a fixed height for the log text area
        self.log_text = ctk.CTkTextbox(log_frame, wrap="word", height=200)
        self.log_text.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")
        
        # Configure row weights to make the log area expand properly
        parent.grid_rowconfigure(row, weight=1)
        
    def browse_directory(self):
        """Open file dialog to choose source directory"""
        directory = filedialog.askdirectory()
        if directory:
            self.selected_directory.set(directory)
            self.populate_subdirectories(directory)
    
    def browse_output_dir(self):
        """Open file dialog to choose output directory"""
        directory = filedialog.askdirectory()
        if directory:
            self.output_directory.set(directory)
    
    def toggle_output_dir(self):
        """Enable/disable custom output directory"""
        state = "normal" if self.custom_output_dir_var.get() else "disabled"
        for widget in self.output_dir_frame.winfo_children():
            widget.configure(state=state)
    
    def populate_subdirectories(self, directory):
        """Populate the subdirectory list"""
        # Clear existing checkboxes
        for widget in self.subdir_scrollable.winfo_children():
            widget.destroy()
        
        # Find all subdirectories
        self.subdirs = []
        self.subdir_vars = {}
        
        try:
            # Start with the root directory itself
            root_var = BooleanVar(value=True)
            root_check = ctk.CTkCheckBox(
                self.subdir_scrollable,
                text=os.path.basename(directory) + " (root)",
                variable=root_var
            )
            root_check.pack(anchor="w", padx=5, pady=2)
            self.subdir_vars[directory] = root_var
            self.subdirs.append(directory)
            
            # Reset selected subdirectories list when changing directory
            self.selected_subdirs = []
            
            # Add subdirectories
            for root, dirs, _ in os.walk(directory):
                for dir_name in dirs:
                    full_path = os.path.join(root, dir_name)
                    rel_path = os.path.relpath(full_path, directory)
                    
                    # Check if this directory contains video files
                    has_videos = False
                    for _, _, files in os.walk(full_path):
                        if any(file.lower().endswith(('.mkv', '.mp4', '.avi', '.mov')) for file in files):
                            has_videos = True
                            break
                    
                    if has_videos:
                        var = BooleanVar(value=True)
                        check = ctk.CTkCheckBox(
                            self.subdir_scrollable,
                            text=rel_path,
                            variable=var
                        )
                        check.pack(anchor="w", padx=5, pady=2)
                        self.subdir_vars[full_path] = var
                        self.subdirs.append(full_path)
        
        except Exception as e:
            self.log_message(f"ERROR: Failed to scan subdirectories: {str(e)}")
        
        # Show the subdirectory frame if we have subdirectories
        if self.subdirs:
            self.subdir_frame.grid()
            # Log a message to prompt for applying selection
            self.log_message("Use 'Apply Selection' to confirm directory selection.")
        else:
            self.subdir_frame.grid_remove()
            self.log_message("No subdirectories with video files found.")
    
    def select_all_subdirs(self):
        """Select all subdirectories"""
        for var in self.subdir_vars.values():
            var.set(True)
    
    def deselect_all_subdirs(self):
        """Deselect all subdirectories"""
        for var in self.subdir_vars.values():
            var.set(False)
    
    def apply_selection(self):
        """Apply subdirectory selection"""
        selected_dirs = self.get_selected_subdirs()
        if not selected_dirs:
            messagebox.showwarning("Warning", "No directories selected.")
            return
        
        self.log_message(f"Selection applied: {len(selected_dirs)} directories selected.")
        
        # List selected directories (limited to first 5 to avoid cluttering the log)
        if len(selected_dirs) <= 5:
            self.log_message(f"Selected directories: {', '.join(os.path.basename(d) for d in selected_dirs)}")
        else:
            dir_names = [os.path.basename(d) for d in selected_dirs[:5]]
            self.log_message(f"Selected directories (showing 5 of {len(selected_dirs)}): {', '.join(dir_names)}...")
        
        # Store selected subdirectories for extraction
        self.selected_subdirs = selected_dirs
        
        # Visual feedback - change apply button color briefly
        original_color = self.apply_btn.cget("fg_color")
        self.apply_btn.configure(fg_color="#4BB543")  # Bright green
        self.after(500, lambda: self.apply_btn.configure(fg_color=original_color))
    
    def get_selected_subdirs(self) -> List[str]:
        """Get list of selected subdirectories"""
        return [subdir for subdir, var in self.subdir_vars.items() if var.get()]
    
    def on_language_change(self, selection):
        """Handle language selection change"""
        if selection == "und":
            # Enable custom language entry if "Undefined" is selected
            self.custom_language_var.set("")
        else:
            # Display the language code
            self.custom_language_var.set(selection)
    
    def get_selected_language(self) -> str:
        """Get the selected language code"""
        # Use custom language if provided, otherwise use selected language
        custom = self.custom_language_var.get().strip()
        if custom:
            return custom.lower()
        return self.language_var.get()
    
    def log_message(self, message):
        """Add a message to the log text area"""
        # Add timestamp
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_message = f"[{timestamp}] {message}\n"
        
        # Insert at the end and scroll to show the latest message
        self.log_text.insert("end", full_message)
        self.log_text.see("end")
    
    def update_progress(self, current, total):
        """Update the progress bar and labels"""
        if total > 0:
            progress = current / total
            self.progress_bar.set(progress)
            
            # Update progress label
            self.progress_label.configure(text=f"Processing: {current}/{total} ({progress:.1%})")
            
            # Update time information
            if self.start_time is not None:
                elapsed = time.time() - self.start_time
                elapsed_str = str(timedelta(seconds=int(elapsed))).split('.')[0]
                
                # Calculate estimated remaining time
                if current > 0:
                    remaining = (elapsed / current) * (total - current)
                    remaining_str = str(timedelta(seconds=int(remaining))).split('.')[0]
                else:
                    remaining_str = "--:--"
                
                self.time_label.configure(text=f"Elapsed: {elapsed_str} | Remaining: {remaining_str}")
    
    def start_extraction(self):
        """Start the subtitle extraction process"""
        # Make sure apply selection was used or fallback to all subdirectories
        if not hasattr(self, 'selected_subdirs') or not self.selected_subdirs:
            # Get selected directories from checkboxes
            selected_dirs = self.get_selected_subdirs()
            if not selected_dirs:
                messagebox.showwarning("Warning", "Please select at least one directory and click 'Apply Selection'.")
                return
            
            # Ask for confirmation if Apply Selection wasn't used
            confirm = messagebox.askyesno(
                "Confirm Selection", 
                "You haven't clicked 'Apply Selection'. Do you want to proceed with the currently selected directories?"
            )
            if not confirm:
                return
                
            self.selected_subdirs = selected_dirs
        
        # Get selected language
        language = self.get_selected_language()
        if not language:
            messagebox.showwarning("Warning", "Please select or enter a language code.")
            return
        
        # Get selected format
        output_format = self.format_var.get()
        
        # Get output directory
        output_dir = None
        if self.custom_output_dir_var.get():
            output_dir = self.output_directory.get()
            if not output_dir:
                messagebox.showwarning("Warning", "Please select an output directory.")
                return
        
        # Clear log before starting new extraction
        self.log_text.delete("1.0", "end")
        self.log_message(f"Starting extraction for {len(self.selected_subdirs)} directories, language: {language}, format: {output_format}")
        
        # Update UI state
        self.extract_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        
        # Reset progress
        self.extractor.reset()
        self.progress_bar.set(0)
        self.progress_label.configure(text="Scanning files...")
        self.time_label.configure(text="Elapsed: 00:00 | Remaining: --:--")
        self.start_time = time.time()
        
        # Start extraction in a thread
        self.extraction_thread = threading.Thread(
            target=self.run_extraction,
            args=(self.selected_subdirs, language, output_format, output_dir)
        )
        self.extraction_thread.daemon = True
        self.extraction_thread.start()
    
    def run_extraction(self, directories, language, output_format, output_dir=None):
        """Run the extraction process in a background thread"""
        try:
            # Get all video files
            video_files = []
            for directory in directories:
                video_files.extend(self.extractor.get_video_files(directory))
            
            if not video_files:
                self.log_message("No video files found in the selected directories.")
                self.finish_extraction()
                return
            
            self.log_message(f"Found {len(video_files)} video files.")
            
            # Scan for subtitle streams
            subtitle_streams = []
            for video_file in video_files:
                if self.extractor.cancel_flag:
                    break
                    
                streams = self.extractor.get_subtitle_streams(video_file)
                matching_streams = [
                    {'video_file': video_file, 'stream': stream}
                    for stream in streams
                    if stream['language'] == language
                ]
                
                subtitle_streams.extend(matching_streams)
                
                # Update progress (based on scan progress)
                self.update_ui_progress(
                    len(subtitle_streams),
                    len(video_files),
                    f"Scanning: {os.path.basename(video_file)}"
                )
            
            # Update total streams count
            self.extractor.total_streams = len(subtitle_streams)
            if subtitle_streams:
                self.log_message(f"Found {len(subtitle_streams)} subtitle streams matching language '{language}'.")
            else:
                self.log_message(f"No subtitle streams found matching language '{language}'.")
                self.finish_extraction()
                return
            
            # Extract subtitles
            self.extractor.processed_streams = 0
            for item in subtitle_streams:
                if self.extractor.cancel_flag:
                    break
                
                video_file = item['video_file']
                stream = item['stream']
                
                # Extract subtitle
                self.extractor.extract_subtitle(
                    video_file=video_file,
                    stream_index=stream['index'],
                    language=language,
                    output_format=output_format,
                    output_dir=output_dir,
                    overwrite=self.overwrite_var.get()
                )
                
                # Update progress
                self.extractor.processed_streams += 1
                self.update_ui_progress(
                    self.extractor.processed_streams,
                    self.extractor.total_streams,
                    f"Extracting: {os.path.basename(video_file)}"
                )
            
            # Finished
            if self.extractor.cancel_flag:
                self.log_message("Extraction cancelled by user.")
            else:
                self.log_message("Extraction process completed.")
                
        except Exception as e:
            self.log_message(f"ERROR: An unexpected error occurred: {str(e)}")
        finally:
            self.finish_extraction()
    
    def update_ui_progress(self, current, total, message="Processing"):
        """Update UI elements from the worker thread"""
        self.after(0, lambda: self.update_progress(current, total))
        self.after(0, lambda: self.progress_label.configure(text=message))
    
    def cancel_extraction(self):
        """Cancel the extraction process"""
        if self.extraction_thread and self.extraction_thread.is_alive():
            self.extractor.cancel()
            self.cancel_btn.configure(state="disabled")
            self.progress_label.configure(text="Cancelling...")
        else:
            self.log_message("No extraction process running.")
    
    def finish_extraction(self):
        """Cleanup after extraction is finished"""
        # Update UI state
        self.extract_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        
        # Final progress update
        if self.extractor.total_streams > 0:
            self.progress_bar.set(1.0)  # Set to 100%
            self.progress_label.configure(text="Completed")
            
            # Display completion statistics
            streams_processed = self.extractor.processed_streams
            total_streams = self.extractor.total_streams
            successful = self.extractor.successful_extractions
            failed = self.extractor.failed_extractions
            skipped = self.extractor.skipped_extractions
            
            if self.extractor.cancel_flag:
                self.log_message(f"Extraction cancelled after processing {streams_processed}/{total_streams} subtitle streams.")
            else:
                if streams_processed > 0:
                    self.log_message("=" * 40)
                    self.log_message(f"EXTRACTION SUMMARY:")
                    self.log_message(f"Processed: {streams_processed}/{total_streams} subtitle streams")
                    self.log_message(f"Successful: {successful} | Failed: {failed} | Skipped: {skipped}")
                    
                    if hasattr(self, 'start_time'):
                        elapsed = time.time() - self.start_time
                        elapsed_str = str(timedelta(seconds=int(elapsed))).split('.')[0]
                        self.log_message(f"Total time: {elapsed_str}")
                    self.log_message("=" * 40)
                    
                    # Show completion notification
                    messagebox.showinfo(
                        "Extraction Complete", 
                        f"Processed {streams_processed} subtitle streams\n\n"
                        f"✓ Successful: {successful}\n"
                        f"✗ Failed: {failed}\n"
                        f"⟳ Skipped: {skipped}"
                    )
        else:
            self.progress_bar.set(0)  # Reset to 0
            self.progress_label.configure(text="Ready")

if __name__ == "__main__":
    app = SubtitleExtractorGUI()
    app.mainloop()