import tkinter as tk
from tkinter import messagebox, filedialog
from graphviz import Digraph
import os

# Sample VBA code for demonstration (replace with actual extraction logic)
sample_vba_code = """
Sub Main()
    ' This subroutine initializes the program.
    Initialize()
    
    Dim x As Integer  ' Variable declaration
    
    ' Perform data processing
    If Condition1 Then
        ProcessData()
    ElseIf Condition2 Then
        ProcessAlternateData()
    Else
        HandleError()
    End If
    
    ' Clean up and end program
    Cleanup()
End Sub
"""

# Function to parse VBA code
def parse_vba_code(vba_code):
    tokens = []
    lines = vba_code.splitlines()
    for line in lines:
        line = line.strip()
        if line.startswith('Sub '):
            tokens.append(('Sub', line, 'Start of program'))
        elif line.startswith('End Sub'):
            tokens.append(('End Sub', line, 'End of program'))
        elif line.startswith('If '):
            tokens.append(('If', line, 'Condition check'))
        elif line.startswith('ElseIf '):
            tokens.append(('ElseIf', line, 'Alternate condition check'))
        elif line.startswith('Else'):
            tokens.append(('Else', line, 'Else condition'))
        elif line.startswith('End If'):
            tokens.append(('End If', line, 'End of condition'))
        elif line.startswith('For '):
            tokens.append(('For', line, 'Loop start'))
        elif line.startswith('Next '):
            tokens.append(('Next', line, 'Loop end'))
        elif line.startswith('Dim ') or line.startswith('Public ') or line.startswith('Private '):
            tokens.append(('Variable Declaration', line, 'Variable declaration'))
        else:
            tokens.append(('Statement', line, 'Statement'))
    return tokens

# Function to generate the flowchart
def generate_flowchart(tokens):
    dot = Digraph()
    counter = 0
    previous_node = None

    for token in tokens:
        node_label = token[1]

        if token[0] == 'Sub':
            dot.node(f'node{counter}', node_label, shape='oval')
            if previous_node is not None:
                dot.edge(f'node{previous_node}', f'node{counter}')
            previous_node = counter
            counter += 1
        elif token[0] == 'End Sub':
            dot.node(f'node{counter}', node_label, shape='oval')
            if previous_node is not None:
                dot.edge(f'node{previous_node}', f'node{counter}')
            previous_node = None
            counter += 1
        elif token[0] == 'If':
            dot.node(f'node{counter}', node_label, shape='diamond')
            if previous_node is not None:
                dot.edge(f'node{previous_node}', f'node{counter}')
            previous_node = counter
            counter += 1
        elif token[0] == 'ElseIf':
            dot.node(f'node{counter}', node_label, shape='diamond')
            if previous_node is not None:
                dot.edge(f'node{previous_node}', f'node{counter}')
            previous_node = counter
            counter += 1
        elif token[0] == 'Else':
            dot.node(f'node{counter}', node_label, shape='diamond')
            if previous_node is not None:
                dot.edge(f'node{previous_node}', f'node{counter}')
            previous_node = counter
            counter += 1
        elif token[0] == 'End If':
            dot.node(f'node{counter}', node_label, shape='diamond')
            if previous_node is not None:
                dot.edge(f'node{previous_node}', f'node{counter}')
            previous_node = counter
            counter += 1
        elif token[0] == 'For':
            dot.node(f'node{counter}', node_label, shape='parallelogram')
            if previous_node is not None:
                dot.edge(f'node{previous_node}', f'node{counter}')
            previous_node = counter
            counter += 1
        elif token[0] == 'Next':
            dot.node(f'node{counter}', node_label, shape='parallelogram')
            if previous_node is not None:
                dot.edge(f'node{previous_node}', f'node{counter}')
            previous_node = counter
            counter += 1
        elif token[0] == 'Statement' or token[0] == 'Variable Declaration':
            if node_label:
                dot.node(f'node{counter}', node_label, shape='box')
                if previous_node is not None:
                    dot.edge(f'node{previous_node}', f'node{counter}')
                previous_node = counter
                counter += 1

    return dot

# Function to open file dialog and select an Excel file
def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xls;*.xlsx;*.xlsm")])
    if file_path:
        process_excel_file(file_path)

# Function to process the selected Excel file
def process_excel_file(file_path):
    # TODO: Implement logic to extract VBA code from the selected Excel file
    tokens = parse_vba_code(sample_vba_code)  # Replace with actual extraction logic
    graph = generate_flowchart(tokens)
    graph.render('flowchart', format='png', cleanup=True)
    os.system('start flowchart.png')  # Automatically open the PNG file in the default image viewer

# Function to go back to the previous page
def back_to_dashboard():
    root.destroy()  # Close the current window
    # Add logic to open the dashboard window here if needed

# Create the tkinter application
root = tk.Tk()
root.title("VBA Macro Flowchart Generator")
root.geometry("1200x900")
root.configure(bg='#5c5b8a')  # Neonish background color

# Description label
description = tk.Label(root, text="This application generates a flowchart from VBA code to visualize program logic.", 
                        font=("Helvetica", 16), bg='#5c5b8a', fg='white')
description.pack(pady=20)

# Button to select an Excel file
btn_generate = tk.Button(root, text="Select Excel File", command=select_excel_file, 
                          font=("Helvetica", 14), bg='#0080ff', fg='white', activebackground='#0059b3')
btn_generate.pack(pady=10)

# Button to go back to the dashboard
btn_back = tk.Button(root, text="Back to Dashboard", command=back_to_dashboard, 
                     font=("Helvetica", 14), bg='#ff5722', fg='white', activebackground='#e64a19')
btn_back.pack(pady=10)

root.mainloop()
