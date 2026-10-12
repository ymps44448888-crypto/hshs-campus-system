import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta

def get_news_html():
    gas_url = "https://script.google.com/macros/s/AKfycbzsvPLtNRUDqpdtIiMt541RKDNWaKZ_iE1r_9BMBz2sBCW6JGunFhQCXpqvBNLAY1wsIA/exec"

    try:
        response = requests.get(gas_url, timeout=20)
        response.encoding = 'utf-8'

        if response.text.startswith("Error:"):
            return f"<p style='color:red;'>Google 中繼站錯誤：{response.text}</p>"

        soup = BeautifulSoup(response.text, 'html.parser')
        target_table = soup.find('table', class_="listTB")

        now = datetime.now()
        current_time_str = now.strftime("%H:%M")
        this_hour_str = now.strftime("%H:00")
        next_hour_str = (now + timedelta(hours=1)).strftime("%H:00")

        news_items = []
        threshold_date = now - timedelta(days=20)
        c = 0
        if target_table:
            rows = target_table.find_all('tr')[1:]
            for row in rows:
                date_td = row.find('td', attrs={'data-th': '日期'})
                date_str = date_td.get_text(strip=True) if date_td else ""
                try:
                    if datetime.strptime(date_str, "%Y-%m-%d") < threshold_date:
                        c += 1
                        if c > 2:
                            break
                except:
                    continue

                mtitle_div = row.find('div', class_='mtitle')
                if mtitle_div:
                    a_tag = mtitle_div.find('a')
                    if a_tag:
                        title = a_tag.get_text(strip=True)
                        link = a_tag.get('href')
                        if link.startswith('p/'):
                            link = "https://www.hshs.ntpc.edu.tw/" + link
                        news_items.append({"date": date_str, "title": title, "link": link})

        content_html = ""
        if not news_items:
            content_html = "<div style='text-align: center; padding: 30px; color: #718096;'>目前沒有 20 天內的訊息</div>"
        else:
            for item in news_items:
                content_html += f"""
                <div style="padding: 15px; border-bottom: 1px dashed #cbd5e0; transition: background-color 0.2s;">
                    <span style="color: #666; font-size: 0.85rem;">[{item['date']}]</span><br>
                    <a href="{item['link']}" target="_blank" style="text-decoration: none; color: #003a7c; font-weight: bold; line-height: 1.5;">{item['title']}</a>
                </div>"""

        html_template = f"""
        <!DOCTYPE html>
        <html lang="zh-TW">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>海山高中班聯會 校網20天內重大公告</title>
            <link rel="stylesheet" href="/static/css/style.css">
            <style>
                .update-info {{ font-size: 0.9rem; color: #555; margin-top: 10px; }}
                h1 {{ margin: 0; font-size: 22px; }}
                strong {{ color: #003a7c; }}
                .news-item-hover:hover {{ background-color: #f7fafc; }}
            </style>
        </head>
        <body>
            <div class="header-box">
                <h1>校網20天內 重要訊息</h1>
                <div class="update-info">
                    現在時間：<strong>{current_time_str}</strong>，
                    {this_hour_str} 已更新，
                    下次更新時間：<strong>{next_hour_str}</strong>
                </div>
            </div>

            <div class="context-card" style="margin-bottom: 25px;">
                <div class="form-group" style="margin-bottom: 0;">
                    <h3>系統說明</h3>
                    <p style="color:#555; padding-left: 15px;">
                        此系統透過 Google 中繼站更新，自動搜尋校網 <strong>[重大公告]</strong> 中，20天內的資訊。<br>
                        若您想了解更多，可按連結跳轉至該公告。
                    </p>
                </div>
            </div>

            <div class="context-card">
                <div class="form-group">
                    <h3>重大公告</h3>
                </div>

                {content_html}

                <div class="footer-flex">
                    <a href="/"><-回首頁</a>
                    <span>by.海山高中班聯會製作</span>
                </div>
            </div>
        </body>
        </html>
        """
        return html_template

    except Exception as e:
        return f"""
        <div style='padding:20px; color:red; border:1px solid red;'>
            <h3>連線中繼站出錯</h3>
            <p>錯誤訊息：{str(e)}</p>
            <p>請檢查 GAS 網址是否正確，且部署時是否有設定為「所有人 (Anyone)」存取。</p>
        </div>
        """

if __name__ == "__main__":
    print(get_news_html())