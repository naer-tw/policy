"""④ 橫向溢出（政策頁）：學校午餐專法懶人包在 320/390/768/1024/1440 寬不得橫向溢出。

修法前：390px 時 `document.documentElement.scrollWidth=543`（比報告數字精確吻合），
根因是三個 `<table class="sr-only">` 在 table-layout:auto 下被內容最小寬度蓋過
`width:1px` 宣告。修法：`table.sr-only{table-layout:fixed}` + `.wrap` 加
`overflow-x:hidden` 防線。
"""
import pytest

WIDTHS = [320, 390, 768, 1024, 1440]
PATH = "/school-lunch/2026-07-10-school-lunch-lazypack.html"


@pytest.mark.parametrize("width", WIDTHS)
def test_no_horizontal_overflow(page, local_site, width):
    page.set_viewport_size({"width": width, "height": 900})
    page.goto(local_site + PATH)
    page.wait_for_timeout(150)

    scroll_width = page.evaluate("document.documentElement.scrollWidth")
    inner_width = page.evaluate("window.innerWidth")

    assert scroll_width <= inner_width, (
        f"{PATH} 在 {width}px 寬度發生橫向溢出："
        f"scrollWidth={scroll_width} > innerWidth={inner_width}"
    )
