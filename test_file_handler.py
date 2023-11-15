"""
File for testing file_handler.py

To pass all tests, you must select
the file /My_xlsx1.xlsx and /My_docx1.docx
on the desktop, which will be created
automatically to automate testing
"""
from dataclasses import dataclass
from unittest import TestCase, main
from file_handler import FileHandler as file_handler
from test_main_class import CommonMethodsTesting as main_class


@dataclass
class Message:
    """
    Сontains message constants
    for output to the user
    """
    WRONG_TYPE_MESSAGE = 'Параметр file_types имеет неправильный тип данных'
    EMPTY_PARAMETER_MESSAGE = 'Параметр file_types не дожен быть пустым'


class TestFileHandler(TestCase):
    """
    Сlass for testing the TestFileHandler class
    Args:
        TestCase (class): inherits from the TestCase class
    """
    pass


class TestGetFilePath(TestCase):
    """
    Class to test get_file_path method
    """
    def test_open_word_file(self) -> None:
        """
        Test to test whether a file is opened correctly.
        For the test to pass, you need to open
        the My_docx1.docx file on your desktop!
        """
        file_path = main_class.create_word_file()
        self.assertEqual(file_handler.get_file_path(
            (("Word file", "*.docx"),)
        ),
            file_path
        )
        main_class.remove_file(file_path)

    def test_open_excel_file(self) -> None:
        """
        Test to test whether a file is opened correctly.
        For the test to pass, you need to open
        the My_xlsx1.xlsx file on your desktop!
        """
        file_path = main_class.create_excel_file()
        self.assertEqual(file_handler.get_file_path(
            (("Excel file", "*.xlsx"),)
        ),
            file_path
        )
        main_class.remove_file(file_path)

    def test_pass_wrong_parameter_list(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.get_file_path(["Excel file", "*.xlsx"])
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_list2(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.get_file_path(["Excel file"])
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_int(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.get_file_path(23)
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_dict(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.get_file_path({23: 'fds'})
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_float(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.get_file_path(56.4)
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_bool1(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.get_file_path(True)
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_bool2(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.get_file_path(False)
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_empty_tuple(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.get_file_path(())
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_tuple_with_empty_str(self) -> None:
        with self.assertRaises(ValueError) as err:
            file_handler.get_file_path(
                (("", ""),
                 ("", ""),)
            )
        self.assertEqual(Message.EMPTY_PARAMETER_MESSAGE, err.exception.args[0])


class TestCheckFileTypes(TestFileHandler):
    def test_open_word_file(self) -> None:
        self.assertEqual(file_handler.check_file_types(
            (("Word file", "*.docx"),)
        ),
            None
        )

    def test_open_excel_file(self) -> None:
        self.assertEqual(file_handler.check_file_types(
            (("Excel file", "*.xlsx"),)
        ),
            None
        )

    def test_pass_wrong_parameter_list(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.check_file_types(["Excel file", "*.xlsx"])
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_list2(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.check_file_types(["Excel file"])
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_int(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.check_file_types(23)
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_dict(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.check_file_types({23: 'fds'})
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_float(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.check_file_types(56.4)
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_bool1(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.check_file_types(True)
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_bool2(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.check_file_types(False)
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_empty_tuple(self) -> None:
        with self.assertRaises(TypeError) as err:
            file_handler.check_file_types(())
        self.assertEqual(Message.WRONG_TYPE_MESSAGE, err.exception.args[0])

    def test_pass_wrong_parameter_tuple_with_empty_str(self) -> None:
        with self.assertRaises(ValueError) as err:
            file_handler.check_file_types(
                (("", ""),
                 ("", ""),)
            )
        self.assertEqual(Message.EMPTY_PARAMETER_MESSAGE, err.exception.args[0])


class TestDisplayFileMessage(TestFileHandler):
    """
    Class for testing the "display_file_message" method
    """
    def test_correct_parameters(self) -> None:
        self.assertEqual(file_handler.display_file_message("str1", "str2"), None)

    def test_with_empty_parameters(self) -> None:
        self.assertEqual(file_handler.display_file_message(), None)


if __name__ == '__main__':
    main()
