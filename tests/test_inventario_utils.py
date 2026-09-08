from inventario_utils import find_sheet


def test_exact_name_is_found():
    assert find_sheet(['RESUMEN', 'TD', 'Stock Libre'], 'Stock Libre') == 'Stock Libre'


def test_uppercase_name_is_found():
    assert find_sheet(['RESUMEN', 'TD', 'STOCK LIBRE'], 'Stock Libre') == 'STOCK LIBRE'


def test_surrounding_spaces_are_ignored():
    assert find_sheet([' stock libre '], 'Stock Libre') == ' stock libre '


def test_missing_sheet_returns_none():
    assert find_sheet(['RESUMEN', 'TD'], 'Stock Libre') is None
