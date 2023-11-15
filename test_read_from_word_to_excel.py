"""
File for testing read_from_word_to_excel.py.

To pass all tests, you must select
the file /My_xlsx1.xlsx and /My_docx1.docx
on the desktop, which will be created
automatically to automate testing
"""
from docx import Document
from unittest import TestCase, main
from test_main_class import CommonMethodsTesting as main_class
from read_from_word_to_excel import WordToExcel as word_to_excel


class TestWordToExcel(TestCase):
    """
    Class for testing the WordToExcel class
    """
    pass


class TestMain(TestWordToExcel):
    """
    Class for testing the main method
    """
    def test_normal_execution(self):
        word_file_path, excel_file_path = main_class.creare_word_and_excel_file()

        self.assertEqual(word_to_excel().main(), None)

        main_class.remove_word_and_excel_file(
            word_file_path,
            excel_file_path
        )


class TestGetDataFromWordFile(TestWordToExcel):
    """
    Class for testing the get_data_from_word_file method
    """
    def test_normal_execution(self):
        word_file_path, excel_file_path = main_class.creare_word_and_excel_file()

        self.assertEqual(word_to_excel().get_data_from_word_file(), None)

        main_class.remove_word_and_excel_file(
            word_file_path,
            excel_file_path
        )


class TestCheckingWordParametrType(TestWordToExcel):
    """
    Class for testing the checking_word_parameter_type method
    """
    def test_standard_load_file(self) -> None:
        word_file_path = main_class.create_word_file()
        word_doc = Document(word_file_path)

        self.assertEqual(word_to_excel().checking_word_parameter_type(word_doc), None)

        main_class.remove_file(word_file_path)

    def test_passing_parameters_of_different_type_str(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().checking_word_parameter_type("simple text")
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])

    def test_passing_parameters_of_different_type_int(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().checking_word_parameter_type(12)
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])

    def test_passing_parameters_of_different_type_float(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().checking_word_parameter_type(34.5)
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])

    def test_passing_parameters_of_different_type_list(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().checking_word_parameter_type(["simple", 34])
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])

    def test_passing_parameters_of_different_type_tuple(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().checking_word_parameter_type((45, 6))
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])

    def test_passing_parameters_of_different_set(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().checking_word_parameter_type({'43', 67})
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])


class TestProcessingReceivedData(TestWordToExcel):
    """
    Class for testing the processing_received_data method
    """
    def test_standard_load_file(self) -> None:
        word_file_path, excel_file_path = main_class.creare_word_and_excel_file()
        word_doc = Document(word_file_path)

        self.assertEqual(word_to_excel().processing_received_data(word_doc), None)

        main_class.remove_word_and_excel_file(
            word_file_path,
            excel_file_path
        )

    def test_passing_parameters_of_different_type_str(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().processing_received_data("simple text")
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])

    def test_passing_parameters_of_different_type_int(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().processing_received_data(23)
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])

    def test_passing_parameters_of_different_type_float(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().processing_received_data(67.8)
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])

    def test_passing_parameters_of_different_type_dict(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().processing_received_data({('sdf'): 'dict'})
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])

    def test_passing_parameters_of_different_type_set(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().processing_received_data({67.8, 'set'})
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])

    def test_passing_parameters_of_different_type_tuple(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().processing_received_data((34, 'tuple'))
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])

    def test_passing_parameters_of_different_type_list(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().processing_received_data([456, 45.6])
        self.assertEqual("Тип передаваемого параметра не 'Document'", err.exception.args[0])


class TectGetExcelParameters(TestWordToExcel):
    """
    Class for testing the get_excel_parametrs method
    """
    DATA_TYPE_ERROR_MESSAGE = "Должна передаваться строка, содержащая путь к файлу"

    def test_wrong_data_type_int(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().get_excel_parameters(45)
        self.assertEqual(
            self.DATA_TYPE_ERROR_MESSAGE,
            err.exception.args[0]
        )

    def test_wrong_data_type_float(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().get_excel_parameters(4.4)
        self.assertEqual(
            self.DATA_TYPE_ERROR_MESSAGE,
            err.exception.args[0]
        )

    def test_wrong_data_type_list(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().get_excel_parameters([1, 3])
        self.assertEqual(
            self.DATA_TYPE_ERROR_MESSAGE,
            err.exception.args[0]
        )

    def test_wrong_data_type_tuple(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().get_excel_parameters((78, 'test'))
        self.assertEqual(
            self.DATA_TYPE_ERROR_MESSAGE,
            err.exception.args[0]
        )

    def test_wrong_data_type_dict(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().get_excel_parameters({78: 'test'})
        self.assertEqual(
            self.DATA_TYPE_ERROR_MESSAGE,
            err.exception.args[0]
        )

    def test_wrong_data_type_set(self) -> None:
        with self.assertRaises(TypeError) as err:
            word_to_excel().get_excel_parameters({1, 3})
        self.assertEqual(
            self.DATA_TYPE_ERROR_MESSAGE,
            err.exception.args[0]
        )


class TestCheckFilePathType(TestWordToExcel):
    """
    Class for testing the check_file_path_type method
    """
    def test_run_without_errors(self):
        self.assertEqual(word_to_excel.check_file_path_type("text"), None)

    INCORRECT_TYPE_MESSAGE = "Должна передаваться строка, содержащая путь к файлу"

    def test_print_wrong_type_error_int(self):
        with self.assertRaises(TypeError) as err:
            word_to_excel().check_file_path_type(12)
        self.assertEqual(self.INCORRECT_TYPE_MESSAGE, err.exception.args[0])

    def test_print_wrong_type_error_float(self):
        with self.assertRaises(TypeError) as err:
            word_to_excel().check_file_path_type(45.6)
        self.assertEqual(self.INCORRECT_TYPE_MESSAGE, err.exception.args[0])

    def test_print_wrong_type_error_list(self):
        with self.assertRaises(TypeError) as err:
            word_to_excel().check_file_path_type(['test', 45])
        self.assertEqual(self.INCORRECT_TYPE_MESSAGE, err.exception.args[0])

    def test_print_wrong_type_error_tuple(self):
        with self.assertRaises(TypeError) as err:
            word_to_excel().check_file_path_type(('t', 45.5))
        self.assertEqual(self.INCORRECT_TYPE_MESSAGE, err.exception.args[0])

    def test_print_wrong_type_error_dict(self):
        with self.assertRaises(TypeError) as err:
            word_to_excel().check_file_path_type({'q1': 'tests'})
        self.assertEqual(self.INCORRECT_TYPE_MESSAGE, err.exception.args[0])


class TestDisplayTypeErrorMessageWd(TestWordToExcel):
    """
    Class for testing the display_type_error_message_Wd method
    """
    def test_standard_execution(self):
        self.assertEqual(word_to_excel.display_type_error_message_Wd(), None)


class TestDisplayAnyErrorMessageWd(TestWordToExcel):
    """
    Class for testing the display_any_error_message_Wd method
    """
    def test_standard_execution(self):
        self.assertEqual(word_to_excel.display_any_error_message_Wd("test"), None)

    INCORRECT_TYPE_MESSAGE = "Должна передаваться строка, содержащая текст ошибки"

    def test_print_wrong_type_error_int(self):
        with self.assertRaises(TypeError) as err:
            word_to_excel.display_any_error_message_Wd(23)
        self.assertEqual(self.INCORRECT_TYPE_MESSAGE, err.exception.args[0])

    def test_print_wrong_type_error_dict(self):
        with self.assertRaises(TypeError) as err:
            word_to_excel.display_any_error_message_Wd({34: 48.34})
        self.assertEqual(self.INCORRECT_TYPE_MESSAGE, err.exception.args[0])


class TestDisplayTypeErrorMessageEx(TestWordToExcel):
    """
    Class for testing the display_type_error_message_Ex method
    """
    def test_standard_execution(self):
        self.assertEqual(word_to_excel.display_type_error_message_Ex("error"), None)

    def test_print_wrong_type_error_int(self):
        with self.assertRaises(TypeError) as err:
            word_to_excel.display_type_error_message_Ex(23)
        self.assertEqual(
            "Должна передаваться строка, содержащая текст ошибки",
            err.exception.args[0]
        )


class TestDisplayAnyErrorMessage_Ex(TestWordToExcel):
    """
    Class for testing the display_any_error_message_Ex method
    """
    def test_standard_execution(self):
        self.assertEqual(word_to_excel.display_any_error_message_Ex("error"), None)

    def test_print_wrong_type_error_int(self):
        with self.assertRaises(TypeError) as err:
            word_to_excel.display_any_error_message_Ex(23)
        self.assertEqual(
            "Должна передаваться строка, содержащая текст ошибки",
            err.exception.args[0]
        )


class TestCheckErrorMessage(TestWordToExcel):
    """
    Class for testing the check_error_message method
    """
    def test_standard_execution(self):
        self.assertEqual(word_to_excel.check_error_message("error"), None)

    def test_print_wrong_type_error_dict(self):
        with self.assertRaises(TypeError) as err:
            word_to_excel.check_error_message({45: 'rt'})
        self.assertEqual(
            "Должна передаваться строка, содержащая текст ошибки",
            err.exception.args[0]
        )


class TestDisplaySuccessMessage(TestWordToExcel):
    """
    Class for testing the display_success_message method
    """
    def test_standard_execution(self):
        self.assertEqual(word_to_excel.display_success_message(), None)


if __name__ == "__main__":
    main()
