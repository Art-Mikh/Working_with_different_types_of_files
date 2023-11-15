"""
File for displaying the main
interface of the program
"""
from tkinter import Tk, Button, Grid
from read_from_word_to_excel import WordToExcel as word_to_excel


class MainWindow:
    """
    Class for displaying the main screen of the program
    Class entry point: main()
    """
    def main(self) -> None:
        """
        Method to call main screen output
        """
        self.open_main_window()

    def open_main_window(self) -> None:
        """
        Method for creating and initializing the main window
        """
        main_window = Tk()
        screen_width = int(main_window.winfo_screenwidth() / 3)
        screen_height = int(main_window.winfo_screenheight() / 3)

        # Screen
        main_window.geometry(f"{screen_width}x{screen_height}")
        main_window.configure(bg='#304A69')  # Background color
        main_window.title("Приложение для файловых операций")

        # Specify Grid
        Grid.rowconfigure(main_window, 0, weight=1)
        Grid.columnconfigure(main_window, 0, weight=1)

        # Button
        word_to_excel_write_button = Button(
            main_window,
            text='Записать данные из Word в Excel',
            command=self.perform_read_from_word_to_excel,
            background="#7F97B4",
            activebackground="#2F3D4E"
        )
        word_to_excel_write_button.pack()  # display a button in the program window
        # Set grid
        word_to_excel_write_button.grid(row=0, column=0, sticky="NSEW")

        # Show window
        main_window.mainloop()

    def perform_read_from_word_to_excel(self) -> None:
        """
        Button method word_to_excel_write_button
        for calls to write data from Word to Excel
        """
        word_to_excel().main()
