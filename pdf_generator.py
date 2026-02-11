"""
Генерация PDF из HTML-презентации через Playwright WebKit (как Safari).
"""
from pathlib import Path
from typing import Optional, Union


def export_html_to_pdf(html_path: Union[str, Path], pdf_path: Optional[Union[str, Path]] = None) -> Path:
    """
    Конвертирует HTML в PDF через WebKit (Safari).

    :param html_path: путь к HTML-файлу
    :param pdf_path: путь для PDF (по умолчанию — рядом с HTML, расширение .pdf)
    :return: путь к созданному PDF
    """
    html_path = Path(html_path).resolve()
    if not html_path.exists():
        raise FileNotFoundError(f"HTML не найден: {html_path}")

    if pdf_path is None:
        pdf_path = html_path.with_suffix(".pdf")
    else:
        pdf_path = Path(pdf_path).resolve()

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise ImportError(
            "Установите Playwright: python3 -m pip install playwright && python3 -m playwright install chromium"
        )

    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(f"file://{html_path}", wait_until="networkidle")
        # Ждём загрузки шрифтов
        page.evaluate("document.fonts.ready")
        # Ждём отрисовки графиков ECharts (canvas)
        page.wait_for_timeout(2500)
        # Печать: режим @media print
        page.emulate_media(media="print")
        page.wait_for_timeout(500)
        # Ресайз графиков под print-размеры
        page.evaluate("""
          (function() {
            if (typeof echarts !== 'undefined') {
              for (const id of ['chart', 'pieChart']) {
                const dom = document.getElementById(id);
                if (dom) {
                  const instance = echarts.getInstanceByDom(dom);
                  if (instance) instance.resize();
                }
              }
            }
          })();
        """)
        page.wait_for_timeout(300)
        page.pdf(
            path=str(pdf_path),
            format="A4",
            landscape=True,
            print_background=True,
            margin={"top": "0", "right": "0", "bottom": "0", "left": "0"},
        )
        browser.close()

    return pdf_path
