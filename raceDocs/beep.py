import tkinter as tk
import winsound

# Function to handle key press events
def on_key_press(event):
    if event.keysym == 'space':
        winsound.Beep(2000, 1000)  # Frequency: 1000 Hz, Duration: 200 ms

# Create the main window
root = tk.Tk()
root.title("Spacebar Beep App")
root.geometry("300x200")

# Create a label to indicate the instruction
label = tk.Label(root, text="Press the spacebar to hear a beep")
label.pack(expand=True)

# Bind the key press event to the on_key_press function
root.bind('<KeyPress>', on_key_press)

# Start the Tkinter event loop
root.mainloop()