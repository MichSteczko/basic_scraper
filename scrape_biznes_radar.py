from typing import List, Dict
from requests import Session
from bs4 import BeautifulSoup
import numpy as np
from pandas import DataFrame, ExcelWriter, concat
from pathlib import Path

def table_hook(response, *args, **kwargs):
    soup = BeautifulSoup(response.content, 'html.parser')
    table = soup.find('table', {'class': 'qTableFull'})
    response._content = table.encode("utf-8")
    return response

def get_session():
    return Session()    

def get_response(*args, **kwargs):
    if "hooks" in kwargs:
        with kwargs["session"].get(kwargs["link"], hooks={'response': kwargs["hooks"]}) as response:
            return response.content
    else:
        with kwargs["session"].get(kwargs["link"]) as response:
            return response.content

def handle_table_elemenst(element):
    try:
        handled_element = element.find("a").text
         
    except AttributeError:
        handled_element = element.text    
       
    return handled_element

def process_column(column_tag):
    try:
        return [handle_table_elemenst(row) for row in column_tag.find_all('td')]
    except:
        return None

def get_dataframe(data, columns):
    return DataFrame(data, columns = columns).fillna(value=np.nan).dropna().reset_index(drop=True)

def get_types(response):
    types_container = response.find("div", {"class": "tools"})
    type_links = types_container.find_all("a")
    type_names = [type_link.find("span").text for type_link in type_links]
    return type_names

def get_category_link(category, url, filtering_types):
    if "kwartalne" in filtering_types:
        return f"{url},Q,{category},2,2,0"
    else:
        return f"{url},0,{category},2,2"

def get_categories(session: Session, url: str) -> List[Dict]:
    response = get_response(session=session, link=url)
    categories = []
    soup = BeautifulSoup(response, 'html.parser')
    filtering_types = get_types(soup)
    select_element = soup.find('select', {'name': 'field'})  # Replace 'select_name' with the actual name of the select element
    for category in select_element.find_all('option'):
        categories.append({f"name": category.text, "link": get_category_link(category["value"], url, filtering_types)})
    return categories

def get_category_dataframe(session: Session, category_map: dict, zone: str, dataframe_cols: list) -> DataFrame:
    dataframe_body = {}
    category_response_bytes = get_response(session=session, link=category_map["link"], hooks=table_hook)
    category_response = BeautifulSoup(category_response_bytes, 'html.parser')
    rows = category_response.find_all('tr')
    processed_rows = [process_column(row)  for row in rows]
    processed_rows_filtered = [val for val in processed_rows if len(val) == 7]
    table_matrix = np.matrix(processed_rows_filtered)
    for index, col in enumerate(dataframe_cols[:-2]):
        dataframe_body.update({col: np.ravel(table_matrix[:, index])})
    rows_num = len(processed_rows_filtered)
    dataframe_body.update({dataframe_cols[-2]: np.array([category_map["name"] for _ in range(rows_num)])})
    dataframe_body.update({dataframe_cols[-1]: np.array([zone for _ in range(rows_num)])})
    category_dataframe = get_dataframe(dataframe_body, columns=dataframe_cols)
    return category_dataframe

def get_dataframe_columns():
    return ["Profil", "Raport", "Wartość", "r/r", "k/k", "r/r>~sektor*", "k/k>~sektor*", "Pozycja", "Sektor"]

def get_sector_links_map() -> Dict[str, str]:
    sectors_map = []
    path_to_sector_links_files = Path("./", "sector_links")
    for sector_links_file in path_to_sector_links_files.iterdir():
        sector_links = []
        with open(sector_links_file, "r", encoding="utf-8") as links_file:
            for link in links_file:
                sector_links.append(link.replace("\n", ""))
        sectors_map.append({"Sector": sector_links_file.name[:-4], "Links": sector_links})
    return sectors_map

def process_biznes_radar_response(sector: str, links: list) -> List[DataFrame]:
    radar_dataframes = []
    session = get_session()
    cols = get_dataframe_columns()

    with session as s:
        for main_link in links:
            categories = get_categories(s, main_link)
            category_dataframes = []
            for category in categories:
                category_dataframes.append(get_category_dataframe(s, category, sector, cols))
            radar_df = concat(category_dataframes).reset_index(drop=True)
            radar_dataframes.append(radar_df)
    return radar_dataframes

def get_full_response_to_save():
    sectors_map = get_sector_links_map()
    sectors_dataframes = []
    multi_ps = []
    multi_bs = []
    multi_cf = []
    for sector_map in sectors_map:
        sectors_dataframes.append(process_biznes_radar_response(sector_map["Sector"], sector_map["Links"]))
    for sector_dataframe in sectors_dataframes:
            pl, bs, cf = sector_dataframe
            multi_ps.append(pl)
            multi_bs.append(bs)
            multi_cf.append(cf)
    
    ps_total = concat(multi_ps)
    bs_total = concat(multi_bs)
    cf_total = concat(multi_cf)
    return [ps_total, bs_total, cf_total]

def save_as_excel(sheet_names: list, dataframes_list: list):
    writer = ExcelWriter("excel_files/biznes_radar.xlsx", engine="xlsxwriter")
    for sheet_name, dataframe in zip(sheet_names, dataframes_list):
        dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
    writer._save()
        

sheet_names = ["P&L", "BS", "CF"]
#main_link = ["https://www.biznesradar.pl/spolki-raporty-finansowe-rachunek-zyskow-i-strat/akcje_gpw", "https://www.biznesradar.pl/spolki-raporty-finansowe-bilans/akcje_gpw", "https://www.biznesradar.pl/spolki-raporty-finansowe-przeplywy-pieniezne/akcje_gpw"]
#bank_link = ["https://www.biznesradar.pl/spolki-raporty-finansowe-rachunek-zyskow-i-strat/branza:banki", "https://www.biznesradar.pl/spolki-raporty-finansowe-bilans/branza:banki", "https://www.biznesradar.pl/spolki-raporty-finansowe-przeplywy-pieniezne/branza:banki"]
dataframes = get_full_response_to_save()
save_as_excel(sheet_names, dataframes)
