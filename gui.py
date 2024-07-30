import subprocess
import tkinter as tk
from tkinter import messagebox, ttk
import re
import os
import sys
import win32com.client  # Ensure pywin32 is installed: pip install pywin32

# Define the path to VBoxManage and log file
VBOXMANAGE_PATH = r"C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"
LOG_FILE_PATH = os.path.join(os.path.expanduser('~'), 'command_log.txt')

def run_as_admin():
    """Relaunch the script with administrative privileges."""
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        script = sys.argv[0]
        command = f"python \"{script}\""
        shell.Run(f"runas /user:Administrator \"{command}\"", 1, True)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to relaunch with admin rights.\n{e}")

def run_command(command):
    """Run a command and return its output."""
    print(f"Executing command: {command}")

    try:
        process = subprocess.Popen(
            command,
            shell=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        stdout, stderr = process.communicate()
        exit_code = process.wait()

        # Print and log the output
        print(f"stdout: {stdout}")
        print(f"stderr: {stderr}")

        # Log the command execution details
        with open(LOG_FILE_PATH, "a") as log_file:
            log_file.write(f"Command executed: {command}\n")
            log_file.write(f"stdout: {stdout}\n")
            log_file.write(f"stderr: {stderr}\n")

        if exit_code != 0:
            raise subprocess.CalledProcessError(exit_code, command, stdout, stderr)

        return stdout

    except subprocess.CalledProcessError as e:
        error_message = (
            f"Command '{e.cmd}' failed with exit status {e.returncode}.\n"
            f"Error Output:\n{e.stderr}\n"
            f"Standard Output:\n{e.stdout}"
        )
        print(error_message)
        with open(LOG_FILE_PATH, "a") as log_file:
            log_file.write(f"Command failed: {e.cmd}\n")
            log_file.write(f"Error Output: {e.stderr}\n")
            log_file.write(f"Standard Output: {e.stdout}\n")
        messagebox.showerror("Error", error_message)
        return ""

    except Exception as e:
        error_message = f"Unexpected error: {str(e)}"
        print(error_message)
        with open(LOG_FILE_PATH, "a") as log_file:
            log_file.write(f"Unexpected error: {str(e)}\n")
        messagebox.showerror("Error", error_message)
        return ""

def list_all_vms():
    """Retrieve the list of all VM names."""
    command = f'"{VBOXMANAGE_PATH}" list vms'
    output = run_command(command)
    if not output.strip():
        messagebox.showwarning("Warning", "No VMs found or there was an issue retrieving VMs.")
    return re.findall(r'"([^"]+)"', output)

def get_network_attachments(vm_name):
    """Retrieve network attachments from VM details."""
    # Quote the VM name to handle spaces
    vm_name_quoted = f'"{vm_name}"'
    command = f'"{VBOXMANAGE_PATH}" showvminfo {vm_name_quoted} --details'
    output = run_command(command)
    attachments = {}
    for line in output.splitlines():
        if 'NIC' in line and 'Cable connected: on' in line:
            match = re.search(r'Attachment: (.*?)(?:,|$)', line)
            if match:
                attachment = match.group(1).strip()
                if attachment not in attachments:
                    attachments[attachment] = []
                attachments[attachment].append(vm_name)
    return attachments

def get_network_interfaces():
    """Retrieve all available network interfaces."""
    interfaces = {}
    for command, interface_type in [
        (f'"{VBOXMANAGE_PATH}" list bridgedifs', "Bridged Interface"),
        (f'"{VBOXMANAGE_PATH}" list hostonlyifs', "Host-Only Interface"),
        (f'"{VBOXMANAGE_PATH}" list intnets', "Internal Network")
    ]:
        output = run_command(command)
        for line in output.splitlines():
            if 'Name:' in line:
                interface = line.split(':', 1)[1].strip()
                interfaces[f"{interface_type} '{interface}'"] = interface
    return interfaces

def update_vm_info():
    """Update and display total unique network interfaces."""
    vms = list_all_vms()
    if not vms:
        messagebox.showinfo("No VMs", "No VMs found.")
        return

    unique_networks = set()
    interface_attachments = {}

    # Collect network attachments for each VM
    for vm in vms:
        attachments = get_network_attachments(vm)
        for interface, connected_vms in attachments.items():
            unique_networks.add(interface)
            if interface not in interface_attachments:
                interface_attachments[interface] = []
            interface_attachments[interface].extend(connected_vms)

    all_interfaces = get_network_interfaces()
    displayed_interfaces = {
        iface for iface in all_interfaces
        if iface in unique_networks
    }

    # Display results
    vm_info_area.config(state=tk.NORMAL)
    vm_info_area.delete(1.0, tk.END)

    total_unique_networks = len(displayed_interfaces)
    vm_info_area.insert(tk.END, f"Total Network Interfaces: {total_unique_networks}\n\n")
    vm_info_area.insert(tk.END, "Network Interfaces:\n")

    for iface in sorted(displayed_interfaces):
        if "Bridged Interface" in iface:
            vm_names = ', '.join(interface_attachments.get(iface, []))
            vm_info_area.insert(tk.END, f"- {iface} (Connected VMs: {vm_names})\n")
        else:
            vm_info_area.insert(tk.END, f"- {iface}\n")

    vm_info_area.config(state=tk.DISABLED)

def show_commands():
    """Display the commands based on user input."""
    vm_name = vm_name_entry.get()
    nic_number = nic_number_entry.get()
    net_name = net_name_entry.get()
    nic_type = nic_type_combobox.get()

    if not vm_name or not nic_number or not net_name or not nic_type:
        messagebox.showerror("Error", "Please fill in all fields.")
        return

    commands = [
        f'"{VBOXMANAGE_PATH}" modifyvm {vm_name} --nic{nic_number} intnet',
        f'"{VBOXMANAGE_PATH}" modifyvm {vm_name} --nictype{nic_number} {nic_type}',
        f'"{VBOXMANAGE_PATH}" modifyvm {vm_name} --intnet{nic_number} {net_name}',
        f'"{VBOXMANAGE_PATH}" modifyvm {vm_name} --cableconnected{nic_number} on'
    ]

    command_output = "\n".join(commands)
    command_display_area.config(state=tk.NORMAL)
    command_display_area.delete(1.0, tk.END)
    command_display_area.insert(tk.END, command_output)
    command_display_area.config(state=tk.DISABLED)

def apply_changes():
    """Apply the commands based on user input."""
    vm_name = vm_name_entry.get()
    nic_number = nic_number_entry.get()
    net_name = net_name_entry.get()
    nic_type = nic_type_combobox.get()

    if not vm_name or not nic_number or not net_name or not nic_type:
        messagebox.showerror("Error", "Please fill in all fields.")
        return

    # Map NIC type descriptions to their actual values
    nic_type_map = {
        'AMD PCNet PCI II': 'Am79C970A',
        'PCNet FAST III (default)': 'Am79C973AMD',
        'Intel PRO/1000 MT Desktop': '82540EM',
        'Intel PRO/1000 T Server': '82543GC',
        'Intel PRO/1000 MT Server': '82545EM',
        'virtio': 'virtio'
    }
    nic_type_value = nic_type_map.get(nic_type, 'virtio')

    commands = [
        f'"{VBOXMANAGE_PATH}" modifyvm {vm_name} --nic{nic_number} intnet',
        f'"{VBOXMANAGE_PATH}" modifyvm {vm_name} --nictype{nic_number} {nic_type_value}',
        f'"{VBOXMANAGE_PATH}" modifyvm {vm_name} --intnet{nic_number} {net_name}',
        f'"{VBOXMANAGE_PATH}" modifyvm {vm_name} --cableconnected{nic_number} on'
    ]

    for command in commands:
        run_command(command)

    messagebox.showinfo("Info", "Changes applied successfully.")

# Set up the GUI
root = tk.Tk()
root.title("VirtualBox VM Network Manager")

frame = ttk.Frame(root, padding="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Create and place GUI components
ttk.Label(frame, text="VM Name:").grid(row=0, column=0, sticky=tk.W)
vm_name_entry = ttk.Entry(frame, width=30)
vm_name_entry.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(frame, text="NIC Number:").grid(row=1, column=0, sticky=tk.W)
nic_number_entry = ttk.Entry(frame, width=30)
nic_number_entry.grid(row=1, column=1, padx=5, pady=5)

ttk.Label(frame, text="Network Name:").grid(row=2, column=0, sticky=tk.W)
net_name_entry = ttk.Entry(frame, width=30)
net_name_entry.grid(row=2, column=1, padx=5, pady=5)

ttk.Label(frame, text="NIC Type:").grid(row=3, column=0, sticky=tk.W)
nic_type_combobox = ttk.Combobox(frame, values=[
    'AMD PCNet PCI II', 'PCNet FAST III (default)', 'Intel PRO/1000 MT Desktop',
    'Intel PRO/1000 T Server', 'Intel PRO/1000 MT Server', 'virtio'
], width=28)
nic_type_combobox.grid(row=3, column=1, padx=5, pady=5)

ttk.Button(frame, text="Show Commands", command=show_commands).grid(row=4, column=0, columnspan=2, pady=10)
ttk.Button(frame, text="Apply Changes", command=apply_changes).grid(row=5, column=0, columnspan=2, pady=10)
ttk.Button(frame, text="Update VM Info", command=update_vm_info).grid(row=6, column=0, columnspan=2, pady=10)

# Text area for displaying commands
command_display_area = tk.Text(frame, height=6, width=60, wrap=tk.WORD, state=tk.DISABLED)
command_display_area.grid(row=7, column=0, columnspan=2, pady=10)

# Text area for displaying VM info
vm_info_area = tk.Text(frame, height=10, width=60, wrap=tk.WORD, state=tk.DISABLED)
vm_info_area.grid(row=8, column=0, columnspan=2, pady=10)

# Run the Tkinter event loop
root.mainloop()
