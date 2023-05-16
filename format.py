import gspread

def format_cells(sheet, oneLessLast, remains, withoutTravel, lastCellCat, lastCellExp):
    sheet.format(f'C2:{oneLessLast}', {
    "backgroundColor": {
    "red": 0.8,
    "green": 0.8,
    "blue": 0.8
    }
    })

    sheet.format('D1', {'text_format': {'bold': True}, 'borders': {'bottom': {'style': 'double'}}})
    sheet.format('C1', {'text_format': {'bold': True}, 'borders': {'bottom': {'style': 'double'}}})
    sheet.format('A1', {'text_format': {'bold': True}, 'borders': {'bottom': {'style': 'double'}}})
    sheet.format('B1', {'text_format': {'bold': True}, 'borders': {'bottom': {'style': 'double'}}})

    sheet.format('A2', {'backgroundColor': {
        'red': 0.6,
        'green': 1,
        'blue': 0.6
    }})

    sheet.format('B2', {'backgroundColor': {
        'red': 0.6,
        'green': 1,
        'blue': 0.6
    }})

    sheet.format('A19', {'backgroundColor': {
        "red": 0.6,
        "green": 1,
        "blue": 0.6
    }
    })

    sheet.format('B19', {'backgroundColor': {
        "red": 0.6,
        "green": 1,
        "blue": 0.6
    }
    })

    sheet.format('A20', {'backgroundColor': {
        "red": 1,
        "green": 0.6,
        "blue": 0.6
    }
    })

    sheet.format('B20', {'backgroundColor': {
        "red": 1,
        "green": 0.6,
        "blue": 0.6
    }
    })

    sheet.format('A21', {'backgroundColor': {
        "red": 1,
        "green": 0.6,
        "blue": 0.6
    }
    })

    sheet.format('B21', {'backgroundColor': {
        "red": 1,
        "green": 0.6,
        "blue": 0.6
    }
    })

    sheet.format('A22', {'backgroundColor': {
        "red": 1,
        "green": 0.6,
        "blue": 0.6
    }
    })

    sheet.format('B22', {'backgroundColor': {
        "red": 1,
        "green": 0.6,
        "blue": 0.6
    }
    })

    if remains < 0:

        sheet.format('A23', {'backgroundColor': {
            "red": 1,
            "green": 0.6,
            "blue": 0.6
        }
        })

        sheet.format('B23', {'backgroundColor': {
            "red": 1,
            "green": 0.6,
            "blue": 0.6
        }
        })

    else:

        sheet.format('A23', {'backgroundColor': {
            "red": 0.2,
            "green": 0.6,
            "blue": 1
        }
        })

        sheet.format('B23', {'backgroundColor': {
            "red": 0.2,
            "green": 0.6,
            "blue": 1
        }
        })

    if withoutTravel < 0:

        sheet.format('C23', {'backgroundColor': {
            "red": 1,
            "green": 0.6,
            "blue": 0.6
        }
        })

        sheet.format('D23', {'backgroundColor': {
            "red": 1,
            "green": 0.6,
            "blue": 0.6
        }
        })

    else:
        sheet.format('C23', {'backgroundColor': {
            "red": 0.2,
            "green": 0.6,
            "blue": 1
        }
        })
        sheet.format('D23', {'backgroundColor': {
            "red": 0.2,
            "green": 0.6,
            "blue": 1
        }
        })

    sheet.format('A23:D23', {'borders': {
        'top': {'style': 'double'}
    }})

    sheet.format(lastCellCat, {'backgroundColor': {
        "red": 1,
        "green": 0.6,
        "blue": 0.6
    }})

    sheet.format(lastCellExp, {'backgroundColor': {
        "red": 1,
        "green": 0.6,
        "blue": 0.6
    }})
