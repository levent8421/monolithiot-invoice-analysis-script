import os

from pallet_sheet import PalletSheet
from xlsx_file_loader import XlsxFileLoader

INPUT_FILE_DIR = 'source'
OUTPUT_FILE_NAME_TEMPLATE = 'target/%s_part.txt'


def group_by_part_no(rows):
    print('Grouping...')
    part_table = {}
    for row in rows:
        part_no = row.part_no
        if part_no in part_table:
            part_table[part_no].append(row)
        else:
            part_table[part_no] = [row]
    print('Group Success!res=', len(part_table))
    return part_table


def reduce_rows(rows, sheet):
    first_row = rows[0]
    part_no = first_row.part_no
    part_name = first_row.part_name
    pcs_sum = 0
    remark_items = []
    pallets = sheet.pallets
    items = [part_no, part_name]
    for row in rows:
        pcs_sum += row.pcs
        remark_items.append('(%s)%s/%s' % (row.pcs, row.pallet_num, pallets))
    remark_str = '; '.join(remark_items)
    items.append(pcs_sum)
    items.append(remark_str)
    return items


def write_headers(f, headers):
    content = '\n'.join([str(header) if header else ' ' for header in headers])
    f.write(content)
    f.write('\n')


def write_result(groups, sheet, filename):
    with open(filename, 'w') as f:
        write_headers(f, sheet.common_headers)
        for part_no, rows in groups.items():
            line_items = reduce_rows(rows, sheet)
            line = '\t'.join([str(item) if item else '' for item in line_items])
            print('LINE[%s]::' % part_no, line)
            f.write(line)
            f.write('\n')


def process_file(filename):
    file_path = os.path.join(INPUT_FILE_DIR, filename)
    loader = XlsxFileLoader(file_path)
    sheet = PalletSheet()
    with loader:
        sheet_name = loader.sheets[0]
        loader.open_sheet(sheet_name)
        sheet.read_headers_from_loader(loader)
        sheet.read_data_rows(loader)
    rows = sheet.rows
    group_res = group_by_part_no(rows)
    out_filename = OUTPUT_FILE_NAME_TEMPLATE % filename
    write_result(group_res, sheet, out_filename)


def run_process():
    files = os.listdir(INPUT_FILE_DIR)
    for file in files:
        process_file(file)


if __name__ == '__main__':
    run_process()
