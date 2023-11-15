import tkinter.filedialog
from tkinter import messagebox


class FileHandler:
    """
    file_handler is a class containing
    the main methods for working with
    files of any type. For example,
    the method is "open file"
    """
    @staticmethod
    def check_file_types(file_types: tuple) -> None:
        WRONG_TYPE_MESSAGE = 'Параметр file_types имеет неправильный тип данных'
        EMPTY_PARAMETER_MESSAGE = 'Параметр file_types не дожен быть пустым'
        if not file_types:
            raise TypeError(WRONG_TYPE_MESSAGE)
        if not isinstance(file_types, tuple):
            raise TypeError(WRONG_TYPE_MESSAGE)
        for item in file_types:
            if not isinstance(item, tuple):
                raise TypeError(WRONG_TYPE_MESSAGE)
            if len(item) != 2:
                raise ValueError(
                    'Параметр file_types должен иметь полное' %
                    'имя расширения файла и сокращенное имя'
                )
            for element in item:
                if not element:
                    raise ValueError(EMPTY_PARAMETER_MESSAGE)

    @classmethod
    def get_file_path(cls, file_types: tuple) -> str:
        """This is a method that calls the file selection dialog box.

        Arg:
            file_types (tuple, optional): Accepts a tuple
            containing a predefined list of openable
            types and a common name of the openable type.
        Example:
            file_types = (("Text file", "*.txt"),
                     ("Image", "*.jpg *.gif *.png"),
                     ("Any", "*"),)

        Return:
            str: full path and name to the selected file
        """
        cls.check_file_types(file_types)
        filename = tkinter.filedialog.askopenfilename(
            title="Open file", initialdir="/",
            filetypes=file_types
        )
        return (filename)

    @classmethod
    def display_file_message(cls, message_title: str = None, message: str = None) -> None:
        """
        error message method
        """
        messagebox.askokcancel(message_title, message)
