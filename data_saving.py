
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import wave
import os
from datetime import datetime
import pyaudio
import threading

class AudioDataCollector:
    def __init__(self, root):
        self.root = root
        root.title("Audio Data Collector")

        self.folder_path = ""
        self.audio_files = []
        self.speaker_vars = []
        self.device_vars = []
        self.excel_path = ""

        self.num_speakers = tk.IntVar(value=2)
        self.play_option = tk.StringVar(value="simultaneous")
        self.save_option = tk.BooleanVar(value=True)
        self.cancel_flag = threading.Event()


        self.ask_excel_file()
        self.create_widgets()

    def ask_excel_file(self):
        messagebox.showinfo("Excel File", "Select the Excel file to use (existing or new).")
        path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not path:
            messagebox.showerror("Error", "Excel file is required to proceed.")
            self.root.destroy()
        self.excel_path = path

        if not os.path.exists(self.excel_path):
            df = pd.DataFrame(columns=["File ID", "Speaker Count", "Playback", "Files Recorded", "Duration Recorded"])
            df.to_excel(self.excel_path, index=False)

    def create_widgets(self):
        row = 0

        # Browse folder
        tk.Button(self.root, text="Browse Folder", command=self.browse_folder).grid(row=row, column=0, padx=5, pady=5)
        self.folder_label = tk.Label(self.root, text="No folder selected", fg="gray")
        self.folder_label.grid(row=row, column=1, columnspan=3, sticky="w")
        row += 1

        # Speaker count
        tk.Label(self.root, text="Number of Speakers:").grid(row=row, column=0, padx=5, pady=5)
        speaker_count_menu = ttk.Combobox(self.root, textvariable=self.num_speakers, values=[1, 2, 3, 4, 5], width=5)
        speaker_count_menu.grid(row=row, column=1)
        speaker_count_menu.bind("<<ComboboxSelected>>", lambda e: self.render_speaker_inputs())
        row += 1

        # Playback options
        tk.Label(self.root, text="Playback:").grid(row=row, column=0, padx=5, pady=5)
        tk.Radiobutton(self.root, text="Simultaneously", variable=self.play_option, value="simultaneous").grid(row=row, column=1, sticky="w")
        tk.Radiobutton(self.root, text="One after other", variable=self.play_option, value="sequential").grid(row=row, column=2, sticky="w")
        row += 1

        # Save option
        tk.Checkbutton(self.root, text="Save to Excel?", variable=self.save_option).grid(row=row, column=0, padx=5, pady=5)
        row += 1

        # Frame for speaker inputs
        self.speaker_frame = tk.Frame(self.root)
        self.speaker_frame.grid(row=row, column=0, columnspan=4, pady=10)
        row += 1

        # Stats label
        self.stats_label = tk.Label(self.root, text="", fg="blue", justify="left")
        self.stats_label.grid(row=row, column=0, columnspan=4)
        row += 1

        # Buttons: Save and Play & Save
        tk.Button(self.root, text="Save Data", command=self.save_to_excel).grid(row=row, column=0, columnspan=2, pady=10, padx=5)
        tk.Button(self.root, text="Play & Save", command=self.play_and_save).grid(row=row, column=2, columnspan=2, pady=10, padx=5)
        tk.Button(self.root, text="Cancel Playback", command=self.cancel_playback).grid(row=row, column=4, pady=10, padx=5)


        # Initial stats
        self.update_stats_label()
        
    def cancel_playback(self):
        self.cancel_flag.set()

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path = folder
            self.audio_files = [f for f in os.listdir(folder)
                                if f.lower().endswith(('.wav', '.mp3', '.aac', '.ogg', '.flac', '.m4a'))]
            self.folder_label.config(text=os.path.basename(folder), fg="black")
            self.render_speaker_inputs()

    def render_speaker_inputs(self):
        for widget in self.speaker_frame.winfo_children():
            widget.destroy()
        self.speaker_vars.clear()
        self.device_vars.clear()

        devices = self.get_playback_devices()
        for i in range(self.num_speakers.get()):
            tk.Label(self.speaker_frame, text=f"Speaker {i+1} File:").grid(row=i, column=0, padx=5, pady=2)
            var = tk.StringVar()
            self.speaker_vars.append(var)
            file_dropdown = ttk.Combobox(self.speaker_frame, textvariable=var, values=self.audio_files, width=50)
            file_dropdown.grid(row=i, column=1)

            tk.Label(self.speaker_frame, text="Playback Device:").grid(row=i, column=2, padx=5)
            dev_var = tk.StringVar()
            self.device_vars.append(dev_var)
            ttk.Combobox(self.speaker_frame, textvariable=dev_var, values=devices, width=40).grid(row=i, column=3)
            
    def get_playback_devices(self):
        pa = pyaudio.PyAudio()
        device_list = []
        for i in range(pa.get_device_count()):
            info = pa.get_device_info_by_index(i)
            if info['maxOutputChannels'] > 0:
                try:
                    # Try to open and immediately close the device to ensure it's usable
                    stream = pa.open(
                        format=pyaudio.paInt16,
                        channels=1,
                        rate=44100,
                        output=True,
                        output_device_index=i
                    )
                    stream.close()
                    device_list.append(f"{i}: {info['name']}")
                except Exception:
                    pass  # Device is not actually available
        pa.terminate()
        return device_list


    def get_duration(self, filename):
        path = os.path.join(self.folder_path, filename)
        try:
            with wave.open(path, 'rb') as wf:
                return wf.getnframes() / wf.getframerate()
        except:
            return 0

    def update_stats_label(self):
        if not os.path.exists(self.excel_path):
            self.stats_label.config(text="No Excel data yet.")
            return
        try:
            df = pd.read_excel(self.excel_path)
            file_count = len(df)
            total_duration = df["Duration Recorded"].sum() if "Duration Recorded" in df else 0
            text = f"📊 Previous Data:\nTotal Files Recorded: {file_count}\nTotal Duration: {total_duration:.2f} s"
            self.stats_label.config(text=text)
        except Exception as e:
            self.stats_label.config(text=f"Error reading Excel file: {e}")

    def save_to_excel(self):
        files = [var.get() for var in self.speaker_vars if var.get()]
        devices = [var.get() for var in self.device_vars if var.get()]
        if len(files) != self.num_speakers.get() or len(devices) != self.num_speakers.get():
            messagebox.showerror("Error", "Please select all speaker files and playback devices.")
            return

        durations = [self.get_duration(f) for f in files]
        if self.play_option.get() == "simultaneous":
            total_duration = max(durations)
        else:
            total_duration = sum(durations)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        df_new = pd.DataFrame([{ 
            "File ID": now,
            **{f"Speaker {i+1} File": files[i] for i in range(len(files))},
            **{f"Speaker {i+1} Device": devices[i] for i in range(len(devices))},
            "Playback": self.play_option.get(),
            "Speaker Count": self.num_speakers.get(),
            "Files Recorded": len(files),
            "Duration Recorded": total_duration
        }])

        if self.save_option.get():
            try:
                if os.path.exists(self.excel_path):
                    df_old = pd.read_excel(self.excel_path)
                    df_all = pd.concat([df_old, df_new], ignore_index=True)
                else:
                    df_all = df_new
                df_all.to_excel(self.excel_path, index=False)
                self.update_stats_label()
                messagebox.showinfo("Saved", f"Data saved to {self.excel_path}\n\n"
                                             f"Files Recorded: {len(files)}\n"
                                             f"Duration: {total_duration:.2f} s")
            except Exception as e:
                messagebox.showerror("Save Error", f"Could not save to Excel: {e}")
        else:
            messagebox.showinfo("Not Saved", "Data not saved to Excel.")

    def play_file(self, filename, device_str):
        try:
            device_index = int(device_str.split(":")[0])
            path = os.path.join(self.folder_path, filename)

            with wave.open(path, 'rb') as wf:
                pa = pyaudio.PyAudio()
                stream = pa.open(
                    format=pa.get_format_from_width(wf.getsampwidth()),
                    channels=wf.getnchannels(),
                    rate=wf.getframerate(),
                    output=True,
                    output_device_index=device_index
                )

                data = wf.readframes(1024)
                while data and not self.cancel_flag.is_set():
                    stream.write(data)
                    data = wf.readframes(1024)

                stream.stop_stream()
                stream.close()
                pa.terminate()
        except Exception as e:
            print(f"Error playing file '{filename}' on device '{device_str}': {e}")



    def play_audio(self):
        # 1) collect selections
        files   = [var.get() for var in self.speaker_vars]
        devices = [var.get() for var in self.device_vars]

        # 2) validate
        if len(files) != self.num_speakers.get() or len(devices) != self.num_speakers.get():
            messagebox.showerror(
                "Error",
                "Please select a file AND a playback device for each speaker before playing."
            )
            return False

        # 3) let the user know
        messagebox.showinfo("Playback", "Starting playback now…")

        self.cancel_flag.clear()

        # 4) play
        if self.play_option.get() == "simultaneous":
            threads = []
            for f, d_str in zip(files, devices):
                t = threading.Thread(target=self.play_file, args=(f, d_str))
                threads.append(t)
                t.start()
            for t in threads:
                t.join()
        else:
            for f, d_str in zip(files, devices):
                self.play_file(f, d_str)

        return True
    def save_session_to_excel(self, files, devices):
        try:
            now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            durations = [self.get_duration(f) for f in files]
            total_duration = max(durations) if self.play_option.get() == "simultaneous" else sum(durations)

            df_new = pd.DataFrame([{ 
                "File ID": now,
                **{f"Speaker {i+1} File": files[i] for i in range(len(files))},
                **{f"Speaker {i+1} Device": devices[i] for i in range(len(devices))},
                "Playback": self.play_option.get(),
                "Speaker Count": self.num_speakers.get(),
                "Files Recorded": len(files),
                "Duration Recorded": total_duration
            }])

            if os.path.exists(self.excel_path):
                df_old = pd.read_excel(self.excel_path)
                df_all = pd.concat([df_old, df_new], ignore_index=True)
            else:
                df_all = df_new
            try:
                with pd.ExcelWriter(self.excel_path, engine='openpyxl', mode='w') as writer:
                    df_all.to_excel(writer, index=False)
            except PermissionError:
                messagebox.showerror("Save Error", "Permission denied. Please close the Excel file if it's open.")
                return
            self.update_stats_label()
            messagebox.showinfo("Saved", f"Data saved to {self.excel_path}\n\n"
                                        f"Files Recorded: {len(files)}\n"
                                        f"Duration: {total_duration:.2f} s")
        except Exception as e:
            messagebox.showerror("Save Error", f"Could not save to Excel: {e}")


    def ask_to_save(self, files, devices):
        def _ask():
            if messagebox.askyesno("Save Data", "Do you want to save this playback session?"):
                self.save_session_to_excel(files, devices)
        self.root.after(0, _ask)

    def play_and_save(self):
        # 1) collect & validate
        files   = [var.get() for var in self.speaker_vars]
        devices = [var.get() for var in self.device_vars]
        if len(files) != self.num_speakers.get() or len(devices) != self.num_speakers.get():
            messagebox.showerror(
                "Error",
                "Please select a file AND a playback device for each speaker before playing."
            )
            return

        # 2) disable buttons so you can't hammer “Play & Save” again
        self.root.nametowidget(".!button2").config(state="disabled")
        self.root.nametowidget(".!button").config(state="disabled")
        
        self.cancel_flag.clear()


        # 3) spawn the background worker
        threading.Thread(
            target=self._play_and_save_worker,
            args=(files, devices),
            daemon=True
        ).start()


    def _play_and_save_worker(self, files, devices):
        # === simultaneous playback ===
        if self.play_option.get() == "simultaneous":
            # 1) open all wave files + streams
            wfs      = []
            pas      = []
            streams  = []
            for filename, device_str in zip(files, devices):
                path = os.path.join(self.folder_path, filename)
                wf   = wave.open(path, 'rb')
                wfs.append(wf)

                pa = pyaudio.PyAudio()
                pas.append(pa)

                device_index = int(device_str.split(":")[0])
                stream = pa.open(
                    format=pa.get_format_from_width(wf.getsampwidth()),
                    channels=wf.getnchannels(),
                    rate=wf.getframerate(),
                    output=True,
                    output_device_index=device_index
                )
                streams.append(stream)

            # 2) read+write in lock‐step until all files are done
            chunk_size = 1024
            while True:
                datas = [wf.readframes(chunk_size) for wf in wfs]
                if all(len(d)==0 for d in datas) or self.cancel_flag.is_set():
                    break
                for data, stream in zip(datas, streams):
                    if data:
                        stream.write(data)  

            # 3) clean up
            for stream, pa, wf in zip(streams, pas, wfs):
                stream.stop_stream()
                stream.close()
                wf.close()
                pa.terminate()

        # === sequential fallback (unchanged) ===
        else:
            for f, d_str in zip(files, devices):
                if self.cancel_flag.is_set():
                    break

                self.play_file(f, d_str)

        # === saving ===
     # Ask whether to save — no matter if playback was cancelled or completed
        self.ask_to_save(files, devices)


        # 4) re-enable buttons on the main thread
        def reenable():
            self.root.nametowidget(".!button2").config(state="normal")
            self.root.nametowidget(".!button").config(state="normal")
        self.root.after(0, reenable)


if __name__ == "__main__":
    root = tk.Tk()
    app = AudioDataCollector(root)
    root.mainloop()
