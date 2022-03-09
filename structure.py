from pathlib import Path

INPUT_FOLDER = Path.cwd() / 'INPUT'
OUTPUT_FOLDER = Path.cwd() / 'OUTPUT'
SRC_FOLDER = Path.cwd() / 'SRC'

COMPANIES = {
    'Fiala': {
        'input_data': INPUT_FOLDER / 'fiala.CSV',
        'src_file_up': SRC_FOLDER / 'jmenny_seznam_2021_10_01 Fiala.xlsx',
        'src_file_loc': SRC_FOLDER / 'Mzdové náklady 2021.xlsx',
        'output_file_up': OUTPUT_FOLDER / 'temp-fiala-up.xlsx',
        'output_file_loc': OUTPUT_FOLDER / 'temp-fiala-loc.xlsx'
    },
    'Bereko': {
        'input_data': INPUT_FOLDER / 'bereko.CSV',
        'src_file_up': SRC_FOLDER / 'jmenny_seznam_2021_10_01 Bereko.xlsx',
        'src_file_loc': SRC_FOLDER / 'Mzdové náklady 2021 Bereko.xlsx',
        'output_file_up': OUTPUT_FOLDER / 'temp-bereko-up.xlsx',
        'output_file_loc': OUTPUT_FOLDER / 'temp-bereko-loc.xlsx'
    }
}
