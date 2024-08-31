from selenium import webdriver
import tkinter as tk
import time
import threading

# GUI for manual user input of partnership type
def show_gui(resume_event):
    root = tk.Tk()

    # Position GUI window at top right
    width = 200
    height = 100
    x = root.winfo_screenwidth() - width
    y = 0
    root.geometry(f"{width}x{height}+{x}+{y}")

    # Create title, label, and button for GUI window
    root.title("Parternership Type")
    prompt = tk.Label(root, text="Please enter this\npartnership's type.")
    prompt.pack(padx=10, pady=5)
    resume = tk.Button(root, text="I'm done", command=lambda: (resume_event.set(), root.withdraw(), root.destroy()))
    resume.pack(pady=10)
    root.mainloop()

def main():
    # Initialize WebDriver
    service = webdriver.ChromeService(executable_path = "../chromedriver-mac-x64/chromedriver")
    driver = webdriver.Chrome(service = service)

    # Navigate to pears.io
    driver.get("https://database-example.com")

    for iter in range(3):
        print("Iteration " + str(iter + 1))
        resume_event = threading.Event() # Create event object for synchronization
        show_gui(resume_event)
        time.sleep(3)
        print("Iteration " + str(iter + 1) + " is completed!")

    # Close WebDriver
    driver.quit()

if __name__ == "__main__":
    main()
