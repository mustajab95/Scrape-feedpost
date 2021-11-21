import time

import pandas

import FileHandling
from selenium.webdriver.chrome.options import Options
from selenium.webdriver import Chrome
from webdriver_manager.chrome import ChromeDriverManager


def fetchYTchannels(link):
    Y_options = Options()
    Y_options.add_argument("--log-level=3")
    Y_options.headless = True
    browser = Chrome(ChromeDriverManager().install(), options=Y_options)
    browser.get(link)
    recent_content = browser.find_elements_by_id('thumbnail')[0:5]
    recent_content = [str(f.get_attribute('href')) for f in recent_content]
    browser.close()
    return recent_content

def feedspot(folder, keyword, link):
    channels_list = []
    browser = Chrome(ChromeDriverManager().install())
    browser.get(link)
    title = browser.find_element_by_css_selector('h1').text.split('Youtube Channels')[0]
    for n in '1234567890':
        title = title.replace(n, '').strip()
    title = ' '.join(title.split(' ')[0:3]) + '...' if len(title) > 20 else title

    h3s = [f.text for f in browser.find_elements_by_css_selector('#fsb > h3 > a.tlink')]
    youtubes = [f.get_attribute('href') for f in browser.find_elements_by_css_selector('#fsb > p.trow-wrap > a.ext')]
    tables = []
    browser.execute_script('d = document.querySelectorAll("#fsb > p.trow-wrap > span.form_sub_wrap > span.vlptext");for(i=0;i<d.length;i++){d[i].click()}')
    while True:
        time.sleep(0.01)
        try:
            tabls = browser.find_elements_by_css_selector('#fsb > p.trow-wrap > span.form_sub_wrap > span.vlp_data > table')
            if len(tabls) == len(h3s):
                break
        except Exception:
            pass

    table_elms = browser.find_elements_by_css_selector('#fsb > p.trow-wrap > span.form_sub_wrap > span.vlp_data > table')
    for i in table_elms:
        trs = i.find_elements_by_css_selector('tbody > tr')
        if len(trs) > 4:
            tables.append([f.find_element_by_css_selector('td:nth-child(2) > a').get_attribute('href') for f in trs])
        else:
            tables.append(fetchYTchannels(youtubes[table_elms.index(i)]))

    rows = [['Channel Name', 'Channel Link', 'Link1', 'Link2', 'Link3', 'Link4', 'Link5']]
    for i in range(len(h3s)):
        row = [h3s[i], youtubes[i], tables[i][0], tables[i][1], tables[i][2], tables[i][3], tables[i][4]]
        channels_list.append(h3s[i])
        rows.append(row)
    path=FileHandling.saveFile(folder, 'Feedspot', 'F_'+keyword, rows)
    browser.close()
    return path

def make_output_file(input_path,output_path):
    try:
        get_sheet_name=pandas.ExcelFile(input_path).sheet_names
        writer = pandas.ExcelWriter(output_path, engine='xlsxwriter')
        writer.save()
        for sheet in get_sheet_name:
            read_excel = pandas.read_excel(input_path , sheet_name=sheet,header=None)
            get_name=read_excel[0].tolist()
            get_excel_link=read_excel[1].tolist()
            counter=0
            path_list=[]
            for name in get_name:
                path=feedspot('Testing', name, get_excel_link[counter])
                counter+=1
                path_list.append(path)
            df_dict={'name':get_name, 'search_link':get_excel_link,'excel_path':path_list }
            df=pandas.DataFrame(df_dict)
            writer=pandas.ExcelWriter(output_path, engine='openpyxl', mode='a')
            df.to_excel(writer,sheet_name=sheet)
            writer.save()
        writer.close()
        return 'done'
    except Exception as err:
        print('FileHandling.py', err)
    return 'Error'


# browser = Chrome(ChromeDriverManager().install())
input_path=r'E:\pycharm_proj\scrape\Files\input\sample feed excel.xlsx'
output_path=r'E:\pycharm_proj\scrape\Files\input\output.xlsx'

make_output_file(input_path,output_path)

#read_excel=pandas.read_excel(input_path)

#feedspot('10Sep21', 'Testing', 'hello','https://blog.feedspot.com/creative_youtube_channels/')
# browser.close()