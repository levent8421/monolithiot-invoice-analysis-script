import os

from pallet_sheet import PalletSheet, Pallet
from xlsx_file_loader import XlsxFileLoader

INPUT_FILE_DIR = 'source/'
OUTPUT_FILE_NAME_TEMPLATE = 'target/%s_pallet.txt'


def group_by_pallet_no(rows):
    print('Grouping by pallet_no...')
    pallet_table = {}
    for row in rows:
        pallet_no = row.pallet_no
        if pallet_no in pallet_table:
            pallet_table[pallet_no].append(row)
        else:
            pallet_table[pallet_no] = [row]
    print('Group success! group size=', len(pallet_table))
    return pallet_table


def convert_to_target_pallets(groups):
    return [Pallet().load_from_rows(groups[pallet_no]) for pallet_no in groups]


def write_result(pallets, sheet, filename):
    print('Writing result...')
    with open(filename, 'w') as f:
        for pallet in pallets:
            line_items = []
            line_items.extend(sheet.common_headers)
            line_items.extend([
                pallet.trace_no, pallet.no, pallet.weight, pallet.size,
            ])
            for payload_item in pallet.payload:
                line_items.extend(payload_item)
            line_str = '\t'.join([str(item) if item else '' for item in line_items])
            print('LINE::', line_str)
            f.write(line_str)
            f.write('\n')
    print('Write success!')


def read_data(loader):
    pallet_sheet = PalletSheet()
    pallet_sheet.read_headers_from_loader(loader)
    pallet_sheet.read_data_rows(loader)
    return pallet_sheet


def process_file(filename):
    file_path = os.path.join(INPUT_FILE_DIR, filename)
    loader = XlsxFileLoader(file_path)
    with loader:
        sheet_name = loader.sheets[0]
        loader.open_sheet(sheet_name)
        sheet_data = read_data(loader)
    group_res = group_by_pallet_no(sheet_data.rows)
    pallets = convert_to_target_pallets(group_res)
    out_filename = OUTPUT_FILE_NAME_TEMPLATE % filename
    write_result(pallets, sheet_data, out_filename)


def run_process():
    files = os.listdir(INPUT_FILE_DIR)
    for file in files:
        if (not file.endswith('.xlsx')) or file.startswith('~'):
            continue
        process_file(file)


if __name__ == '__main__':
    run_process()
