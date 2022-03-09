from os import getenv
from dataclasses import asdict, dataclass
from datetime import datetime, timedelta
from typing import List

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from sp_api.api import Orders
from sp_api.base import Marketplaces, SellingApiException

import const


""" classe para inserção autenticação e definição de métodos 
    que serão utilizados no Google Planilha """

class GoogleSheets:
    def __init__(self, credentials_file: str, sheet_key: str, worksheet_key: str): 
        self.credentials_file = credentials_file
        self.sheet_key = sheet_key
        self.worksheet_key = worksheet_key
        self.scope = ["https://spreadsheets.google.com/feeds"]
        self.sheet_object = self._get_sheet_object()

    
    """método da classe para receber credenciais e abertura do Google Planilhas"""
    
    def _get_sheet_object(self):
        credentials = ServiceAccountCredentials.from_json_keyfile_name(
            self.scope, self.credentials_file
        )
        client = gspread.authorize(credentials)
        return client.open_by_key(self.sheet_key).woksheet(self.worksheet_name)


    """método da classe para escrever cabeçalho na planilha, caso ele não exista""" 
        
    def write_header_if_doesnt_exist(self, columns: List[str]) -> None:
        data = self.sheet_object.get_all_values()
        if not data:
            self.sheet_object.insert_row(columns)


    """método da classe para adicionar linhas"""
    
    def append_row(self, rows: List[str]) -> None:
        last_row_number = len(self.sheet_object.col_values(1)) + 1
        self.sheet_object.insert_rows(rows, last_row_number)


@dataclass
class AmazonOrder:
    order_id: str
    purchase_date: str
    order_total: str 
    payment_method: str 
    marketplace_ide: str 
    shipment_service_level_category: str 
    order_type: str 
    

"""definindo cabeção que será incluído na planilha"""

HEADER = [
    'AmazonOrderId',
    'PurchaseDate',
    'OrderStatus',
    'OrderTotal',
    'PaymentMethod',
    'MarketplaceId',
    'ShipmentServiceLevelCategory',
    'OrderType',
]



"""classe para obter dados da API da Amazon e \
    preencher a planilha no Google Planilhas"""

class AmazonScript: 
    def __init__(self):
        google_sheets = GoogleSheets(
            "keys.json", os.getenv.GOOGLE_SHEETS_ID, os.getenv.GOOGLE_WORKSHEET_NAME
        )
        google_sheets.write_header_if_doesnt_exist(HEADER)
        self.get_orkers_data_and_append_to_gs(google_sheets)

    
    """método para receber as ordens de pedido da 
       plantaforma e adicionar ao Google Planilha"""    
        
    def get_orders_data_and_append_to_gs(self, google_sheets: GoogleSheets) -> None:
        try:
            order_data = self.get_orders_from_sp_api()
            ready_rows= [list(asdict(row).values()) for row in order_data]
            google_sheets.append_row(ready_rows)
        except SellingApiException as e:
            print(f'Error: {e}')
        


    


    
