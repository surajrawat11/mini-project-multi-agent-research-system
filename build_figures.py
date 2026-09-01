"""Generates the figure images embedded in the internship report."""

import os
from PIL import Image, ImageDraw, ImageFont

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
os.makedirs(OUT, exist_ok=True)

W = 1600
SCALE = 2
INK = (25, 25, 25)
GREY = (110, 110, 110)
FILL = (238, 242, 248)
ACCENT = (46, 82, 138)
BAR = (70, 110, 170)
BAR2 = (170, 190, 215)


def font(size, bold=False):
    name = "timesbd.ttf" if bold else "times.ttf"
    for path in (rf"C:\Windows\Fonts\{name}", name):
        try:
            return ImageFont.truetype(path, size)
        except OSError:
            continue
    return ImageFont.load_default()


def canvas(height):
    img = Image.new("RGB", (W, height), "white")
    return img, ImageDraw.Draw(img)


def centre(d, box, text, f, colour=INK):
    x0, y0, x1, y1 = box
    l, t, r, b = d.textbbox((0, 0), text, font=f)
    d.text((x0 + (x1 - x0 - (r - l)) / 2 - l, y0 + (y1 - y0 - (b - t)) / 2 - t), text, font=f, fill=colour)


def box(d, xy, lines, f_head, f_body, fill=FILL, outline=ACCENT):
    x0, y0, x1, y1 = xy
    d.rounded_rectangle(xy, radius=10, fill=fill, outline=outline, width=3)
    n = len(lines)
    h = (y1 - y0) / n
    for i, (text, bold) in enumerate(lines):
        centre(d, (x0, y0 + i * h, x1, y0 + (i + 1) * h), text, f_head if bold else f_body)


def arrow(d, x, y0, y1, w=4):
    d.line([(x, y0), (x, y1 - 16)], fill=ACCENT, width=w)
    d.polygon([(x - 12, y1 - 18), (x + 12, y1 - 18), (x, y1)], fill=ACCENT)


def save(img, name):
    path = os.path.join(OUT, name)
    img.resize((W // SCALE, img.height // SCALE), Image.LANCZOS).save(path, dpi=(300, 300))
    print("wrote", path)


def fig_3_1():
    img, d = canvas(1180)
    fh, fb = font(38, True), font(32)
    layers = [
        ("PRESENTATION LAYER", "Streamlit web interface - topic input, progress, results"),
        ("ORCHESTRATION LAYER", "LangChain - agent lifecycle, chaining, callbacks, retries"),
        ("AGENT LAYER", "Search Agent | Reader Agent | Writer Agent | Critic Agent"),
        ("SERVICE LAYER", "DuckDuckGo Search API | Groq gpt-oss-120b | BeautifulSoup4"),
    ]
    y = 60
    for i, (head, body) in enumerate(layers):
        box(d, (140, y, W - 140, y + 190), [(head, True), (body, False)], fh, fb)
        if i < len(layers) - 1:
            arrow(d, W / 2, y + 190, y + 265)
        y += 265
    save(img, "fig_3_1_architecture.png")


def fig_3_2():
    img, d = canvas(1420)
    fh, fb = font(36, True), font(30)
    stages = [
        ("USER INPUT", "Research topic submitted through the web interface"),
        ("SEARCH AGENT  (7.2 s)", "DuckDuckGo query -> relevance scoring -> top 10 URLs"),
        ("READER AGENT  (11.5 s)", "Fetch pages -> parse HTML -> strip boilerplate -> aggregate"),
        ("WRITER AGENT  (18.3 s)", "Groq LLM synthesis -> structured Markdown report"),
        ("CRITIC AGENT  (12.1 s)", "Five-dimension scoring -> quality score and feedback"),
        ("RESULT DELIVERED  (49.1 s total)", "Report, score and source list shown to the user"),
    ]
    y = 50
    for i, (head, body) in enumerate(stages):
        shade = (226, 234, 245) if i in (0, len(stages) - 1) else FILL
        box(d, (120, y, W - 120, y + 170), [(head, True), (body, False)], fh, fb, fill=shade)
        if i < len(stages) - 1:
            arrow(d, W / 2, y + 170, y + 225)
        y += 225
    save(img, "fig_3_2_pipeline.png")


def fig_3_3():
    img, d = canvas(610)
    fh, fb, fs = font(34, True), font(28), font(26)
    names = ["Search\nAgent", "Reader\nAgent", "Writer\nAgent", "Critic\nAgent"]
    payload = ["10 ranked\nURLs", "8k - 15k words\nof text", "Markdown\nreport", ""]
    x = 60
    bw, gap = 280, 120
    for i, name in enumerate(names):
        d.rounded_rectangle((x, 240, x + bw, 440), radius=10, fill=FILL, outline=ACCENT, width=3)
        for j, line in enumerate(name.split("\n")):
            centre(d, (x, 270 + j * 48, x + bw, 318 + j * 48), line, fh)
        if i < len(names) - 1:
            ax0, ax1 = x + bw, x + bw + gap
            d.line([(ax0 + 10, 340), (ax1 - 26, 340)], fill=ACCENT, width=4)
            d.polygon([(ax1 - 28, 328), (ax1 - 28, 352), (ax1 - 8, 340)], fill=ACCENT)
            for j, line in enumerate(payload[i].split("\n")):
                centre(d, (ax0, 150 + j * 42, ax1, 192 + j * 42), line, fs, GREY)
        x += bw + gap
    centre(d, (0, 60, W, 110), "Data handed from each agent to the next", font(34, True))
    centre(d, (0, 520, W, 570), "State is held in memory for the duration of one request only", fb, GREY)
    save(img, "fig_3_3_dataflow.png")


def fig_4_1():
    img, d = canvas(1250)
    fh, fb = font(34, True), font(30)
    steps = [
        "Receive research topic from the pipeline",
        "Build search query and call the DuckDuckGo API",
        "Score each result:  0.4 x title match  +  0.6 x snippet match",
        "Discard results scoring below the 0.5 threshold",
        "Remove duplicate URLs",
        "Sort by descending relevance and return the top 10",
    ]
    y = 60
    for i, text in enumerate(steps):
        box(d, (150, y, W - 150, y + 130), [(f"{i + 1}.  {text}", False)], fh, fb)
        if i < len(steps) - 1:
            arrow(d, W / 2, y + 130, y + 185)
        y += 185
    save(img, "fig_4_1_search_agent.png")


def fig_4_2():
    img, d = canvas(1250)
    fh, fb = font(34, True), font(30)
    steps = [
        "Take the ranked URL list from the Search Agent",
        "Request each page with a 10 second timeout",
        "Parse the response with BeautifulSoup4",
        "Remove script, style, nav, header and footer elements",
        "Reject any extraction shorter than 100 words",
        "Concatenate the surviving text and record failure counts",
    ]
    y = 60
    for i, text in enumerate(steps):
        box(d, (150, y, W - 150, y + 130), [(f"{i + 1}.  {text}", False)], fh, fb)
        if i < len(steps) - 1:
            arrow(d, W / 2, y + 130, y + 185)
        y += 185
    save(img, "fig_4_2_reader_agent.png")


def fig_4_3():
    img, d = canvas(1080)
    fh, fb = font(34, True), font(28)
    d.rectangle((110, 60, W - 110, 1010), outline=INK, width=4)
    d.rectangle((110, 60, W - 110, 210), fill=(232, 238, 247), outline=INK, width=4)
    centre(d, (110, 60, W - 110, 210), "Multi-Agent Research System", font(46, True))

    d.rectangle((170, 270, W - 170, 400), outline=GREY, width=3)
    centre(d, (190, 270, 1000, 400), "Enter a research topic", fb, GREY)
    d.rounded_rectangle((1130, 295, 1380, 375), radius=8, fill=ACCENT)
    centre(d, (1130, 295, 1380, 375), "Research", font(30, True), "white")

    d.rectangle((170, 450, W - 170, 560), outline=GREY, width=3)
    d.rectangle((173, 453, 950, 557), fill=(196, 214, 240))
    centre(d, (170, 450, W - 170, 560), "Stage 3 of 4  -  Writer Agent  -  31 s elapsed", fb)

    d.rectangle((170, 620, W - 170, 960), outline=GREY, width=3)
    centre(d, (190, 640, 900, 700), "Generated research report", fh)
    for i in range(4):
        d.line([(210, 740 + i * 46), (1100, 740 + i * 46)], fill=(200, 200, 200), width=6)
    d.rounded_rectangle((1170, 660, 1400, 780), radius=8, fill=(232, 240, 232), outline=(90, 140, 90), width=3)
    centre(d, (1170, 660, 1400, 725), "Quality", fb)
    centre(d, (1170, 715, 1400, 780), "8.2 / 10", font(40, True))
    save(img, "fig_4_3_interface.png")


def bar_chart(name, title, labels, series, colours, legend, ymax, ylabel, fmt="{:.1f}"):
    img, d = canvas(900)
    ft, fl, fv = font(38, True), font(28), font(26)
    x0, y0, x1, y1 = 200, 150, W - 120, 720
    centre(d, (0, 45, W, 110), title, ft)
    d.line([(x0, y1), (x1, y1)], fill=INK, width=3)
    d.line([(x0, y0), (x0, y1)], fill=INK, width=3)

    ticks = 5
    for i in range(ticks + 1):
        v = ymax * i / ticks
        y = y1 - (y1 - y0) * i / ticks
        d.line([(x0 - 12, y), (x1, y)], fill=(224, 224, 224), width=2)
        d.text((x0 - 100, y - 16), f"{v:.0f}", font=fv, fill=INK)
    d.text((60, y0 - 60), ylabel, font=fl, fill=INK)

    n, m = len(labels), len(series)
    slot = (x1 - x0) / n
    bw = slot * 0.62 / m
    for i, label in enumerate(labels):
        base = x0 + slot * i + slot * 0.19
        for j, values in enumerate(series):
            h = (y1 - y0) * values[i] / ymax
            bx0 = base + j * bw
            d.rectangle((bx0, y1 - h, bx0 + bw - 6, y1), fill=colours[j])
            centre(d, (bx0, y1 - h - 46, bx0 + bw - 6, y1 - h - 6), fmt.format(values[i]), fv)
        centre(d, (x0 + slot * i, y1 + 14, x0 + slot * (i + 1), y1 + 70), label, fl)

    lx = x0
    for j, text in enumerate(legend):
        d.rectangle((lx, 790, lx + 40, 826), fill=colours[j])
        d.text((lx + 56, 788), text, font=fl, fill=INK)
        lx += 460
    save(img, name)


def fig_5_1():
    bar_chart(
        "fig_5_1_performance.png",
        "Measured Agent Execution Time Against Target",
        ["Search", "Reader", "Writer", "Critic", "Pipeline"],
        [[7.2, 11.5, 18.3, 12.1, 49.1], [10, 15, 20, 15, 60]],
        [BAR, BAR2],
        ["Measured mean (seconds)", "Target ceiling (seconds)"],
        60,
        "Seconds",
    )


def fig_5_2():
    bar_chart(
        "fig_5_2_coverage.png",
        "Statement Coverage by Module Against the 85 Per Cent Target",
        ["Search", "Reader", "Writer", "Critic", "Pipeline", "Overall"],
        [[92, 88, 82, 79, 84, 85]],
        [BAR],
        ["Statement coverage (per cent)"],
        100,
        "Per cent",
        fmt="{:.0f}",
    )


if __name__ == "__main__":
    fig_3_1()
    fig_3_2()
    fig_3_3()
    fig_4_1()
    fig_4_2()
    fig_4_3()
    fig_5_1()
    fig_5_2()
