# -*- coding: utf-8 -*-
import argparse
import json
import re
from pathlib import Path
from playwright.sync_api import sync_playwright


DEFAULT_END_DAY = 20  # PDF 截止日，可直接修改；也可用 --end-day 指定。


def prepare_pdf_employee(employee, end_day):
    from copy import deepcopy
    from datetime import datetime
    result = deepcopy(employee)
    days = []
    for day in result.get('出勤', []):
        date = datetime.strptime(str(day.get('日期', '')), '%Y/%m/%d')
        if date.day <= end_day:
            days.append(day)
    result['出勤'] = days
    required = float(result.get('應休時數', float(result.get('應休天數', 0)) * 8))
    used = sum(float(day.get('計休時數') or 0) for day in days)
    result['應休時數'] = required
    result['已休時數'] = used
    result['未休時數'] = max(0.0, required - used)
    result['系統加班累計'] = sum(float(day.get('系統加班') or 0) for day in days)
    result['遲到累計'] = sum(float(day.get('遲到分鐘') or 0) for day in days)
    return result


def safe_name(value):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', str(value)).strip().rstrip(' .')
    if not name:
        name = '未分類'
    if name.split('.')[0].upper() in {'CON', 'PRN', 'AUX', 'NUL', *{f'COM{i}' for i in range(1,10)}, *{f'LPT{i}' for i in range(1,10)}}:
        name = '_' + name
    return name


def main():
    parser = argparse.ArgumentParser(description='依門別建立資料夾，沿用 HTML 列印樣式輸出員工 PDF')
    parser.add_argument('--html', default='hours.html')
    parser.add_argument('--json', default='員工出勤資料_10月.json')
    parser.add_argument('--limit', type=int, default=5, help='先測試5位；0表示全部')
    parser.add_argument('--month', default='10月')
    parser.add_argument('--end-day', type=int, default=DEFAULT_END_DAY, help='PDF明細及摘要截止日，預設20')
    args = parser.parse_args()
    if not 1 <= args.end_day <= 31:
        parser.error('--end-day 必須介於1與31之間')
    base = Path(__file__).resolve().parent
    html = (base / args.html).read_text(encoding='utf-8-sig')
    data = json.loads((base / args.json).read_text(encoding='utf-8-sig'))
    if not isinstance(data, list) or args.limit < 0:
        raise ValueError('JSON 必須為員工陣列，limit 不可小於0')
    selected = data[:args.limit] if args.limit else data
    html, count = re.subn(r'\bloadData\(\);(?=\s*</script>)', '', html)
    if count != 1:
        raise ValueError('找不到 loadData() 啟動點，請使用原有網頁程式結構')
    out = base / f'PDF_{safe_name(args.month)}_截至{args.end_day:02d}日_{"測試" if args.limit else "全部"}'
    jobs = []
    seen = set()
    for original_employee in selected:
        employee = prepare_pdf_employee(original_employee, args.end_day)
        gate = employee.get('顯示門別') or employee.get('門別') or '未分類'
        filename = safe_name(str(employee['員編']) + '_' + str(employee['姓名'])) + '.pdf'
        target = out / safe_name(gate) / filename
        key = str(target).casefold()
        if key in seen:
            raise ValueError('輸出檔名重複：' + str(target))
        seen.add(key)
        if target.exists():
            raise FileExistsError('請先移走舊PDF再執行：' + str(target))
        jobs.append((employee, target))
    with sync_playwright() as p:
        browser = p.chromium.launch()
        try:
            page = browser.new_page(viewport={'width': 1500, 'height': 1000})
            page.set_content(html, wait_until='load')
            page.emulate_media(media='print')
            for index, (employee, target) in enumerate(jobs, 1):
                target.parent.mkdir(parents=True, exist_ok=True)
                page.evaluate("({employee, month}) => {currentMonth = month; const card = renderAttendance(employee, {forPdf: true}); employeeContainer.replaceChildren(card); loading.style.display = 'none'; noResult.style.display = 'none'; resultInfo.textContent = '';}",
                              {'employee': employee, 'month': args.month})
                page.evaluate('document.fonts.ready')
                page.evaluate('() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))')
                page.pdf(path=str(target), print_background=True,
                         prefer_css_page_size=True, display_header_footer=False, scale=1)
                if target.stat().st_size == 0:
                    raise RuntimeError('PDF為空：' + target.name)
                print(f'[{index}/{len(jobs)}] {target.relative_to(out)}')
        finally:
            browser.close()
    print('完成，輸出位置：' + str(out))


if __name__ == '__main__':
    main()
