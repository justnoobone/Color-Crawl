import requests
from requests.adapters import HTTPAdapter
from lxml import etree
from openpyxl import Workbook
from fake_useragent import UserAgent

# 常量配置
BASE_URL = 'https://www.color-name.com/colors/{color}'  # 修正URL格式
EXCEL_PATH = 'color_links.xlsx'
HEADERS = {'User-Agent': UserAgent().random}

COLORS = [
    'blue', 'teal', 'green', 'yellow', 'orange', 'red', 'pink', 'purple',
    'gray', 'silver', 'white', 'black', 'gold', 'olive', 'khaki', 'beige',
    'brown', 'chocolate', 'maroon', 'indigo', 'navy', 'cyan', 'aqua', 'fuchsia'
]


def get_color_links(color: str, max_retries: int = 3) -> list[str]:
    """根据颜色名称生成动态URL并抓取链接"""
    session = requests.Session()
    session.mount('https://', HTTPAdapter(max_retries=max_retries))

    try:
        # 动态生成URL
        url = BASE_URL.format(color=color.lower())
        response = session.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
        tree = etree.HTML(response.content)


        # 方案1：绝对路径
        links = tree.xpath('/html/body/div[2]/div/ul//a/@href')
        # 方案2：属性过滤（假设父级div有class="main-content"）
        # links = tree.xpath('//div[@class="main-content"]/div/ul//a/@href')

        return links if links else []
    except Exception as e:
        print(f"Error fetching {color}: {e}")
        return []


def save_to_excel(all_links: list[str]) -> None:
    """去重后保存到Excel（优化写入性能）"""
    unique_links = list(dict.fromkeys(all_links))  # 去重保留顺序

    wb = Workbook()
    ws = wb.active
    ws.title = "颜色链接"
    ws.append(['序号', '颜色链接'])

    # 修正逐行写入逻辑
    for idx, link in enumerate(unique_links, start=1):
        ws.append([idx, link])

    wb.save(EXCEL_PATH)
    wb.close()


if __name__ == '__main__':
    all_links = []

    # 遍历所有颜色
    for color in COLORS:
        print(f"正在抓取 {color}...")
        links = get_color_links(color)
        if links:
            all_links.extend(links)

    if all_links:
        save_to_excel(all_links)
        print(f"已保存{len(all_links)}条数据到{EXCEL_PATH}（去重后{len(set(all_links))}条）")
    else:
        print("未获取到有效链接")