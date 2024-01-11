from tkinter import Tk, Button, Label, mainloop

def choose_report(on_report_selected):
    def report_selected(report_number):
        on_report_selected(report_number)

    # Create the main UI window
    root = Tk()
    root.title("Choose Report")

    # Add label
    label = Label(root, text="Select the report to export:")
    label.pack(pady=10)

    # Add buttons for each report
    button_report1 = Button(root, text="Report 1", command=lambda: report_selected(1))
    button_report1.pack(pady=5)

    button_report2 = Button(root, text="Report 2", command=lambda: report_selected(2))
    button_report2.pack(pady=5)

    # Run the UI loop
    root.mainloop()
