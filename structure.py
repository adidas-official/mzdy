from pathlib import Path

INPUT_FOLDER = Path.cwd() / 'INPUT'
OUTPUT_FOLDER = Path.cwd() / 'OUTPUT'
SRC_FOLDER = Path.cwd() / 'SRC'

COMPANIES = {
    'Fiala': {
        'input_data': INPUT_FOLDER / 'fiala.CSV',
        'src_file': SRC_FOLDER / 'jmenny_seznam_2021_10_01 Fiala.xlsx',
        'output_file': OUTPUT_FOLDER / 'temp-fiala.xlsx'
    },
    'Bereko': {
        'input_data': INPUT_FOLDER / 'bereko.CSV',
        'src_file': SRC_FOLDER / 'jmenny_seznam_2021_10_01 Bereko.xlsx',
        'output_file': OUTPUT_FOLDER / 'temp-bereko.xlsx'
    }
}
