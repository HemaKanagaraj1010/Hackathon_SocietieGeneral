import tkinter as tk
import subprocess

def open_automated_vba_macro_documentation():
    subprocess.Popen(["python", "C:\\Users\\Hemalatha\\Downloads\\vba_macro_analyzer (1)\\vba_macro_analyzer\\vba_documentation.docx"])

def open_automated_process_flow_visualization():
    subprocess.Popen(["python", "C:\\Users\\Hemalatha\\Downloads\\vba_macro_analyzer (1)\\vba_macro_analyzer\vba_documentation.docx"])

# Setting up the GUI
root = tk.Tk()
root.title("VBA Macro Analyzer Dashboard")
root.geometry("800x600")
root.configure(bg='#1e1e78')  # Neonish blue color

# Create a title label
title_label = tk.Label(root, text="VBA Macro Analyzer Dashboard", font=("Helvetica", 24, "bold"), bg='#1e1e78', fg='#00ffcc')
title_label.pack(pady=20)

# Description label
description_label = tk.Label(
    root, 
    text=("This dashboard provides easy access to tools for automating the documentation of VBA macros and visualizing their processes. "
          "Select one of the options below to proceed."),
    bg="#1e1e78", 
    fg="#ffffff", 
    font=("Helvetica", 16), 
    wraplength=750, 
    justify="center"
)
description_label.pack(pady=20)

# Create a frame for buttons
frame = tk.Frame(root, bg='#1e1e78')
frame.pack(pady=50)

# Style for buttons
button_style = {
    'font': ("Helvetica", 14),
    'bg': '#00ccff',
    'fg': 'white',
    'activebackground': '#0099cc',
    'relief': 'flat',
    'borderwidth': 1,
    'width': 40,
    'height': 2
}

# Button for Automated VBA Macro Documentation
btn_automated_vba_macro_documentation = tk.Button(
    frame,
    text="Automated VBA Macro Documentation",
    command=open_automated_vba_macro_documentation,
    **button_style
)
btn_automated_vba_macro_documentation.pack(pady=10)

# Button for Automated Process Flow Visualization
btn_automated_process_flow_visualization = tk.Button(
    frame,
    text="Automated Process Flow Visualization",
    command=open_automated_process_flow_visualization,
    **button_style
)
btn_automated_process_flow_visualization.pack(pady=10)

# Run the main loop
root.mainloop()
