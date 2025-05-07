import logging
import tkinter as tk
import tkinter.scrolledtext as st
from threading import Thread

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s - %(message)s',
    handlers=[
        logging.FileHandler("main.log", encoding='utf-8'),
        logging.StreamHandler()
    ],
    force=True
)

import process_emails

last_log = ""

window = tk.Tk()
window.geometry('800x600')
window.title("Email Handling")

def run_script():
    Thread(target=process_emails.main, daemon=True).start()

def read_log():
    global last_log
    try:
        with open('main.log', 'r', encoding='utf-8') as file:
            current_log = file.read()
        if current_log != last_log:
            
            yview = text.yview()
            at_bottom = yview[1] == 1.0  

            text.delete(1.0, tk.END)
            for line in current_log.splitlines():
                if "INFO" in line:
                    text.insert(tk.END, line + "\n", "INFO")
                elif "WARNING" in line:
                    text.insert(tk.END, line + "\n", "WARNING")
                elif "ERROR" in line:
                    text.insert(tk.END, line + "\n", "ERROR")
                elif "DEBUG" in line:
                    text.insert(tk.END, line + "\n", "DEBUG")
                elif "CRITICAL" in line:
                    text.insert(tk.END, line + "\n", "CRITICAL")
                else:
                    text.insert(tk.END, line + "\n")


            if at_bottom:
                text.see(tk.END)

            last_log = current_log
    except Exception as e:
        text.insert(tk.END, f'Error reading log {e}')
        
    window.after(1000, read_log)


runScriptButton = tk.Button(window, text='Run Script', command=run_script)
runScriptButton.pack()

text = st.ScrolledText(window)

text.tag_config("INFO", foreground="green")
text.tag_config("WARNING", foreground="orange")
text.tag_config("ERROR", foreground="red")
text.tag_config("DEBUG", foreground="gray")
text.tag_config("CRITICAL", foreground="magenta")

text.pack(fill=tk.BOTH, expand=True)


read_log()

window.mainloop()