import time
import psutil
import pymem
import rtmidi
import tkinter as tk
from tkinter import ttk, scrolledtext, Toplevel, Label, Menu, Scale
import os
import threading
import queue
import webbrowser

"""
RSTone2MIDI
Author: Marian-Mina Mihai (bboylalu)
Github: https://github.com/bboylalu/RSTone2MIDI
"""

"""
The code was written with the help of Google Gemini, so if you want to throw rocks, please rock on \m/!!!
"""

def get_process_id_by_window_title(window_title):
    """Gets the process ID of a process by its name."""
    try:
        for proc in psutil.process_iter():
            try:
                if proc.name() == "Rocksmith2014.exe":
                    return proc.pid
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass
        return None
    except Exception as e:
        print(f"Error iterating processes: {e}")
        return None

def get_module_base_address(process_handle, module_name):
    """Gets the base address of a module in a process."""
    try:
        modules = process_handle.list_modules()
        for module in modules:
            if module.name == module_name:
                return module.lpBaseOfDll
        return None
    except Exception as e:
        print(f"Error getting base address: {e}")
        return None

def read_memory_with_offsets(process_handle, base_address, base_pointer_offset, offsets):
    """Reads a value from memory, handling null pointers."""
    try:
        base_pointer_address = base_address + base_pointer_offset
        current_address = process_handle.read_int(base_pointer_address)

        if current_address == 0:
            return None  # Return None if base pointer is null

        for offset in offsets[:-1]:
            current_address += offset
            current_address = process_handle.read_int(current_address)
            if current_address == 0:
                return None  # Return None if any pointer in the chain is null

        final_offset = offsets[-1]
        final_address = current_address + final_offset
        value = process_handle.read_int(final_address)
        return value

    except Exception as e:
        print(f"An error occurred in read_memory_with_offsets: {e}")
        return None

def send_midi_control_change(channel, control, value, midi_out):
    """Sends a MIDI Control Change message using rtmidi."""
    try:
        cc_message = [0xB0 | (channel - 1), control, value]
        midi_out.send_message(cc_message)
        #print(f"Sent MIDI CC: Channel {channel}, Control {control}, Value {value}")  # Now handled by the GUI
    except Exception as e:
        print(f"Error sending MIDI CC: {e}")

def send_midi_program_change(channel, program, midi_out):
    """Sends a MIDI program change message using rtmidi."""
    try:
        program_change_message = [0xC0 | (channel - 1), program]
        midi_out.send_message(program_change_message)
        #print(f"Sent MIDI: Channel {channel}, Program {program}")  # Now handled by the GUI
    except Exception as e:
        print(f"Error sending MIDI: {e}")

def is_game_running(window_title):
    """Checks if the game is running."""
    return get_process_id_by_window_title(window_title) is not None

CONFIG_FILE = "RSTone2MIDI_config.txt"

def read_config():
    """Reads MIDI port and message type from config file."""
    try:
        with open(CONFIG_FILE, "r") as f:
            lines = f.readlines()
            if len(lines) >= 2:  # Ensure both port and message type are present
                try:
                    port = int(lines[0].strip())
                    message_type = lines[1].strip().lower()  # Store message type as lowercase
                    return port, message_type
                except ValueError:
                    print(f"Invalid data in {CONFIG_FILE}. Please correct it.")
                    return None, None
            else:
                return None, None  # File is empty or doesn't contain both values
    except FileNotFoundError:
        return None, None  # File doesn't exist

def write_config(port, message_type):
    """Writes MIDI port and message type to config file."""
    try:
        with open(CONFIG_FILE, "w") as f:
            f.write(f"{port}\n{message_type}")  # Write both port and message type
    except Exception as e:
        print(f"Error writing to {CONFIG_FILE}: {e}")

def update_gui_messages(q):
    """Updates the message display in the GUI."""
    try:
        while True:
            message = q.get_nowait()
            message_display.insert(tk.END, message + "\n")
            message_display.see(tk.END)
    except queue.Empty:
        pass
    root.after(100, update_gui_messages, q)  # Correct: root is now in scope

def main_loop(q, window_title, module_name):
    """The main game processing loop, running in a separate thread."""
    
    last_slider_value = 0  # Initialize slider value

    available_ports = rtmidi.MidiOut().get_ports()
    selected_port, selected_message_type = read_config()

    if selected_port is None or selected_message_type is None:
        q.put("Missing port or message type selection. Please configure in the GUI.")
        return  # Exit the thread

    if selected_port not in range(len(available_ports)):
        q.put(f"Port {selected_port} from the config file is not available. Please choose a valid port in the GUI.")
        return

    try:  # Try to open the MIDI port
        midi_out = rtmidi.MidiOut()
        midi_out.open_port(selected_port)
        
        waiting_for_window_message_printed = False  # Flag for "Waiting for song..." message

        while True:
        
            if not is_game_running(window_title):
                slider_value = midi_slider.get()
                if slider_value != last_slider_value:  # Check if slider value changed
                    midi_value = slider_value
                    if selected_message_type == "control change":
                        send_midi_control_change(1, 1, midi_value, midi_out)
                        q.put(f"Sent MIDI Control Change: Channel 1, Control 1, Value {midi_value}")
                    elif selected_message_type == "program change":
                        send_midi_program_change(1, midi_value, midi_out)
                        q.put(f"Sent MIDI Program Change: Channel 1, Program {midi_value}")
                    last_slider_value = slider_value
                    
                if not waiting_for_window_message_printed:
                    q.put(f"Before you launch the game you can test your MIDI connectivity / map your controls with the help of the slider below.\n{window_title} is not running. Waiting...")
                    waiting_for_window_message_printed = True
                time.sleep(1)
                continue
                
            waiting_for_window_message_printed = False # Reset the flag when a valid tone_id is read
            
            # Check Game State
            game_running = is_game_running(window_title)

            # Visual Feedback for Slider State
            if game_running:
                midi_slider.config(troughcolor="indian red")  # Change trough color to gray
                midi_slider.config(state=tk.DISABLED)
            else:
                midi_slider.config(troughcolor="gray50")  # Reset trough color (or set to your default)
                midi_slider.config(state=tk.NORMAL)

            pid = get_process_id_by_window_title(window_title)

            try:
                pm = pymem.Pymem(pid)

                base_address = get_module_base_address(pm, module_name)

                if base_address is None:
                    q.put(f"Module '{module_name}' not found in process {pid}.")
                    continue

                # ***REPLACE THESE WITH YOUR ACTUAL VALUES***
                base_pointer_offset = 0xF5F54C  # Replace with your actual base pointer offset
                offsets = [
                    0x10,
                    0x28,
                    0x38,
                    0x18,
                    0x04,
                    0xBC,
                    0x10
                ]

                last_tone_id = None
                waiting_for_song_message_printed = False

                while is_game_running(window_title):
                    tone_id = read_memory_with_offsets(pm, base_address, base_pointer_offset, offsets)

                    if tone_id is None:
                        if not waiting_for_song_message_printed:
                            q.put("Waiting for song...")
                            waiting_for_song_message_printed = True
                        time.sleep(1)
                        continue

                    waiting_for_song_message_printed = False
                    
                    if tone_id != last_tone_id:
                        if tone_id == 5:
                            cc_value = 0
                        else:
                            cc_value = tone_id

                        if selected_message_type == "control change":
                            send_midi_control_change(1, 1, cc_value, midi_out)
                            q.put(f"Sent MIDI Control Change: Channel 1, Control 1, Value {cc_value}")
                        elif selected_message_type == "program change":
                            send_midi_program_change(1, cc_value, midi_out)
                            q.put(f"Sent MIDI Program Change: Channel 1, Program {cc_value}")
                        last_tone_id = tone_id

                    time.sleep(0.1)

                del pm
                q.put(f"{window_title} closed. Waiting for it to restart...")

            except pymem.exception.ProcessNotFound:
                q.put(f"{window_title} is not running. Waiting...")
                time.sleep(1)
                continue

            except Exception as e:
                q.put(f"An error occurred: {e}")
                break

        del midi_out  # Close the MIDI port when the game closes

    except Exception as e:  # Catch any exceptions during MIDI port opening or usage
        q.put(f"Error with MIDI: {e}")
        return  # Exit the thread if there's an error with MIDI
          
def open_config_window():
    """Opens the MIDI configuration window."""
    config_window = Toplevel(root)  # Create a new top-level window
    config_window.title("MIDI Configuration")

    # Styling for a nicer look (same as in the main window creation)
    config_window.configure(padx=20, pady=20, bg="#EEEEEE")
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('TCombobox', padding=5, fieldbackground="#EEEEEE", background="#EEEEEE")
    style.configure('TButton', padding=5, background="#EEEEEE")
    style.configure('TLabel', background="#EEEEEE")

    port_label = ttk.Label(config_window, text="Select MIDI Port:")
    port_label.pack(pady=(0, 5))

    port_var = tk.StringVar(config_window)
    available_ports = rtmidi.MidiOut().get_ports() # Get ports inside the function
    port_options = [f"{i}: {port_name}" for i, port_name in enumerate(available_ports)]

    if port_options:
        port_dropdown = ttk.Combobox(config_window, textvariable=port_var, values=port_options, state="readonly")
        port_dropdown.current(0)
        port_dropdown.pack(pady=(0, 10))

    message_type_label = ttk.Label(config_window, text="Select Message Type:")
    message_type_label.pack(pady=(10, 5))

    message_type_var = tk.StringVar(config_window)
    message_type_options = ["Program Change", "Control Change"]
    message_type_dropdown = ttk.Combobox(config_window, textvariable=message_type_var, values=message_type_options, state="readonly")
    message_type_dropdown.current(0)
    message_type_dropdown.pack(pady=(0, 10))

    def save_config():
        if port_var.get() and message_type_var.get():
            selected_port = int(port_var.get().split(":")[0])
            selected_message_type = message_type_var.get().lower()
            write_config(selected_port, selected_message_type)
            config_window.destroy()  # Close the config window
            # You might want to restart the main loop thread here to apply the new config
        else:
            print("Missing port or message type selection.")


    select_button = ttk.Button(config_window, text="Save Configuration", command=save_config)
    select_button.pack()

    config_window.transient(root) # Appear on top
    config_window.grab_set() # Make it modal
    
def open_about_window():
    """Opens the 'About' window."""
    about_window = Toplevel(root)
    about_window.title("About RSTone2MIDI")

    about_text_parts = [
        "RSTone2MIDI v1.0\n",
        "Author: Marian-Mina Mihai (bboylalu)\n",
        "GitHub: ",  # Text before the link
        "https://github.com/bboylalu/RSTone2MIDI",  # The actual link
        "This program allows you to send MIDI messages based on the current tone in Rocksmith 2014.\n",
        "Disclaimer:",
        "I am not a professional programmer, just some guy passionate about guitar and programming and an ethical tinkerer.",
        "Therefore I cannot provide professional support and the app is provided as is with no responsibility on my side, or yours.",
        "I made this as a command line script initially, for my own private use, then did my best to make it user-friendly for regular users.",
        "Feel free to contact me if you encounter any issues and I'll do my best to help, depending on my skills and available time.",
        "\nSpecial Thanks:",
        "* RSMods and all their contributors",
        "https://github.com/Lovrom8/RSMods",
        "* RS ASIO",
        "https://github.com/mdias/rs_asio",
        "* Rocksmith Custom Song Toolkit and all their contributors",
        "https://github.com/rscustom/rocksmith-custom-song-toolkit",
        "* CustomsForge and all their contributors",
        "https://customsforge.com/",
        "* Cheat Engine",
        "https://www.cheatengine.org/",
        "* My girlfriend for her patience",
        "* Google Gemini for helping me write the script",
        "\nThe script is written in Python and it uses the following libraries: psutil, pymem, rtmidi, tkinter",
        "and it was packed as an *.exe with pyinstaller",
        "\nDonations are welcome!"
    ]

    # Create Labels for the text parts and the link
    for part in about_text_parts:
        if part.startswith("https://"):  # It's a link
            link_label = Label(about_window, text=part, fg="blue", cursor="hand2")
            link_label.pack(anchor=tk.W)

            # The crucial fix: using a closure
            def create_link_command(link):  # Inner function to create the command
                return lambda event: webbrowser.open_new(link)  # Capture the 'link' value

            link_label.bind("<Button-1>", create_link_command(part))  # Use the closure

        else:
            text_label = Label(about_window, text=part, justify=tk.LEFT)
            text_label.pack(anchor=tk.W)

    about_window.resizable(False, False)
    about_window.transient(root)
    about_window.grab_set()
    
def open_help_link():
    help_link = "https://github.com/bboylalu/RSTone2MIDI"  # Replace with your actual help link
    webbrowser.open_new(help_link)

if __name__ == "__main__":
    window_title = "Rocksmith 2014"
    module_name = "Rocksmith2014.exe"

    available_ports = rtmidi.MidiOut().get_ports()
    selected_port, selected_message_type = read_config()

    if selected_port is None or selected_message_type is None:
        if available_ports:
            root = tk.Tk()
            root.title("MIDI Configuration")

            # Styling for a nicer look
            root.configure(padx=20, pady=20)
            style = ttk.Style()
            style.theme_use('clam')  # Use a modern theme
            style.configure('TCombobox', padding=5)  # Adjust combobox styling
            style.configure('TButton', padding=5)  # Adjust button styling

            # Uniform background (light gray)
            root.configure(bg="#EEEEEE")  # Light gray background for the main window
            style.configure('TCombobox', fieldbackground="#EEEEEE", background="#EEEEEE")  # Light gray for comboboxes
            style.configure('TButton', background="#EEEEEE")  # Light gray for buttons
            style.configure('TLabel', background="#EEEEEE") # Light gray for labels

            port_label = ttk.Label(root, text="Select MIDI Port:")
            port_label.pack(pady=(0, 5))

            port_var = tk.StringVar(root)
            port_options = [f"{i}: {port_name}" for i, port_name in enumerate(available_ports)]

            if port_options:
                port_dropdown = ttk.Combobox(root, textvariable=port_var, values=port_options, state="readonly")
                port_dropdown.current(0)
                port_dropdown.pack(pady=(0, 10))

            message_type_label = ttk.Label(root, text="Select Message Type:")
            message_type_label.pack(pady=(10, 5))

            message_type_var = tk.StringVar(root)
            message_type_options = ["Program Change", "Control Change"]
            message_type_dropdown = ttk.Combobox(root, textvariable=message_type_var, values=message_type_options, state="readonly")
            message_type_dropdown.current(0)
            message_type_dropdown.pack(pady=(0, 10))

            select_button = ttk.Button(root, text="Select", command=lambda: root.destroy())
            select_button.pack()


            root.mainloop()

            if port_var.get() and message_type_var.get():
                selected_port = int(port_var.get().split(":")[0])
                selected_message_type = message_type_var.get().lower()
                write_config(selected_port, selected_message_type)

            else:
                print("Missing port or message type selection.")
                exit()

        else:
            print("No MIDI ports found.")
            exit()
    elif selected_port not in range(len(available_ports)):
        print(f"Port {selected_port} from the config file is not available. Please choose a valid port.")
        selected_port = None
        if available_ports:
            root = tk.Tk()
            root.title("MIDI Configuration")

            port_var = tk.StringVar(root)
            port_options = [f"{i}: {port_name}" for i, port_name in enumerate(available_ports)]

            if port_options:
                port_dropdown = ttk.Combobox(root, textvariable=port_var, values=port_options, state="readonly")
                port_dropdown.current(0)  # Select the first port by default
                port_dropdown.pack(pady=10)

            message_type_var = tk.StringVar(root)
            message_type_options = ["Program Change", "Control Change"]
            message_type_dropdown = ttk.Combobox(root, textvariable=message_type_var, values=message_type_options, state="readonly")
            message_type_dropdown.current(0)
            message_type_dropdown.pack(pady=10)

            select_button = ttk.Button(root, text="Select", command=lambda: root.destroy())
            select_button.pack()

            root.mainloop()

            if port_var.get() and message_type_var.get():
                selected_port = int(port_var.get().split(":")[0])
                selected_message_type = message_type_var.get().lower()
                write_config(selected_port, selected_message_type)

            else:
                print("Missing port or message type selection.")
                exit()

        else:
            print("No MIDI ports found.")
            exit()

    root = tk.Tk()  # Define root *before* using it
    root.title("RSTone2MIDI")
    
    # Create Menubar
    menubar = Menu(root)
    root.config(menu=menubar)
    
    # Create File Menu
    filemenu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=filemenu)

    # Add Config and Exit to File Menu
    filemenu.add_command(label="Settings", command=open_config_window)
    filemenu.add_separator() # Separator line
    filemenu.add_command(label="Exit", command=root.destroy)
    
    # Create Help Menu
    helpmenu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Help", menu=helpmenu)

    helpmenu.add_command(label="How to use", command=open_help_link)
    helpmenu.add_separator() # Separator line
    
    # Add About to Help Menu
    helpmenu.add_command(label="About", command=open_about_window)

    message_display = scrolledtext.ScrolledText(root, wrap=tk.WORD)
    message_display.pack(expand=True, fill=tk.BOTH)

    q = queue.Queue()

    main_thread = threading.Thread(target=main_loop, args=(q, window_title, module_name))
    main_thread.daemon = True
    main_thread.start()

    root.after(100, update_gui_messages, q)
        
    # MIDI Value Slider
    midi_slider = Scale(root, from_=0, to=3, orient=tk.HORIZONTAL, label="Test MIDI (0-3)")
    midi_slider.pack(pady=(10, 0))
    midi_slider.set(0)
    
    root.mainloop()
