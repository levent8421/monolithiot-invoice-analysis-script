INVOICE_NO_POSITION = 'B3'
TRACKING_NO_POSITION = 'D3'
PALLETS_POSITION = 'F3'
CARRIER_POSITION = 'H3'
RECEIPT_POSITION = 'J3'
RECEIPT_ORG_POSITION = 'L3'
CONSIGNEE_POSITION = 'J4'
CONSIGNEE_CONCAT_POSITION = 'L4'
DETAIL_ADDRESS_POSITION = 'B4'
PACKAGE_DATE_POSITION = 'B5'
SHIPMENT_POSITION = 'D5'
EXPECTED_TIME_POSITION = 'F5'
CONSIGNOR_POSITION = 'H5'
SHIPPER_POSITION = 'J5'
CONSIGNOR_CONCAT_POSITION = 'L5'

DATA_ROW_OFFSET = 7

PART_NO_INDEX = 1
PART_NAME_INDEX = 2
PCS_INDEX = 3
REMARK_INDEX = 4
PALLET_NO_INDEX = 5
PALLET_NUM_INDEX = 6
SIZE_INDEX = 7
WEIGHT_INDEX = 8

MAX_ROWS = 10 * 1000


class PalletSheet:
    def __init__(self):
        #  发货单号
        self.invoice_no = None
        # 运单号
        self.tracking_no = None
        # 托数
        self.pallets = 0
        # 承运商
        self.carrier = None
        # 收货地
        self.receipt = None
        # 收货方
        self.receipt_org = None
        # 收货人
        self.consignee = None
        # 收货人联系方式
        self.consignee_concat = None
        # 详细地址
        self.detail_address = None
        # 装货日期
        self.package_date = None
        # 发货日期
        self.shipment = None
        # 预期送达时间
        self.expected_time = None
        # 发货方
        self.shipper = None
        # 发货人
        self.consignor = None
        # 发货人联系方式
        self.consignor_concat = None
        # 数据行
        self.rows = []

    def read_headers_from_loader(self, loader):
        self.invoice_no = loader.read_cell(INVOICE_NO_POSITION)
        self.tracking_no = loader.read_cell(TRACKING_NO_POSITION)
        self.pallets = loader.read_cell(PALLETS_POSITION)
        self.carrier = loader.read_cell(CARRIER_POSITION)
        self.receipt = loader.read_cell(RECEIPT_POSITION)
        self.receipt_org = loader.read_cell(RECEIPT_ORG_POSITION)
        self.consignee = loader.read_cell(CONSIGNEE_POSITION)
        self.consignee_concat = loader.read_cell(CONSIGNEE_CONCAT_POSITION)
        self.detail_address = loader.read_cell(DETAIL_ADDRESS_POSITION)
        self.package_date = loader.read_cell(PACKAGE_DATE_POSITION)
        self.shipment = loader.read_cell(SHIPMENT_POSITION)
        self.expected_time = loader.read_cell(EXPECTED_TIME_POSITION)
        self.shipper = loader.read_cell(SHIPPER_POSITION)
        self.consignor = loader.read_cell(CONSIGNOR_POSITION)
        self.consignor_concat = loader.read_cell(CONSIGNOR_CONCAT_POSITION)

    def read_data_rows(self, loader):
        self.rows = read_rows(loader)

    @property
    def common_headers(self):
        sheet = self
        return [sheet.invoice_no, sheet.tracking_no, sheet.pallets, sheet.carrier, sheet.receipt,
                sheet.receipt_org, sheet.consignee, sheet.consignee_concat, sheet.detail_address,
                sheet.package_date, sheet.shipment, sheet.expected_time, sheet.shipper, sheet.consignor,
                sheet.consignor_concat, ]


class Pallet:
    def __init__(self):
        self.trace_no = None
        self.no = None
        self.weight = None
        self.size = None
        self.payload = []

    def load_from_rows(self, rows):
        first_row = rows[0]
        self.trace_no = first_row.pallet_no
        self.no = first_row.pallet_num
        self.weight = first_row.weight
        self.size = first_row.size
        for row in rows:
            payload_item = (row.part_no, row.part_name, row.pcs, row.remark)
            self.payload.append(payload_item)
        return self


class DataRow:
    def __init__(self):
        self.index = None
        self.part_no = None
        self.part_name = None
        self.pcs = None
        self.remark = None
        self.pallet_no = None
        self.pallet_num = None
        self.size = None
        self.weight = None

    def load_from_sheet_row(self, row):
        self.part_no = row[PART_NO_INDEX].value
        self.part_name = row[PART_NAME_INDEX].value
        self.pcs = row[PCS_INDEX].value
        self.remark = row[REMARK_INDEX].value
        self.pallet_no = row[PALLET_NO_INDEX].value
        self.pallet_num = row[PALLET_NUM_INDEX].value
        self.size = row[SIZE_INDEX].value
        self.weight = row[WEIGHT_INDEX].value


def read_rows(loader):
    sheet = loader.sheet
    row_num = DATA_ROW_OFFSET
    res_rows = []
    while row_num < MAX_ROWS:
        row = sheet[row_num]
        data_row = DataRow()
        data_row.load_from_sheet_row(row)
        data_row.index = row_num
        if not data_row.pallet_no:
            print('Loaded rows:', row_num - DATA_ROW_OFFSET)
            break
        res_rows.append(data_row)
        row_num += 1
    return res_rows
