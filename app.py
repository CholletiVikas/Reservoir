from flask import Flask, render_template, request, send_from_directory
from openpyxl import load_workbook
import re
import os

app = Flask(__name__)

OUTPUT_FOLDER = "output"

os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def parse_messages(text):

    data = {}
    date = ""

    date_match = re.search(r'(\d{2}[.-]\d{2}[.-]\d{4})', text)

    if date_match:
        date = date_match.group(1).replace("-", ".")

    # Gandiramaram
    g = re.search(
        r'Gandiramaram.*?Present Level\s*:\s*([\d.]+).*?Present Storage\s*:\s*([\d.]+)',
        text,
        re.I | re.S
    )

    if g:
        data["Gandiramaram"] = [
            float(g.group(1)),
            float(g.group(2))
        ]

    # Bommakur
    b = re.search(
        r'Bommakur.*?Present Level\s*:\s*([\d.]+).*?Present Storage\s*:\s*([\d.]+)',
        text,
        re.I | re.S
    )

    if b:
        data["Bommakur"] = [
            float(b.group(1)),
            float(b.group(2))
        ]

    # Ashwaraopally
    a = re.search(
        r'Ashwaraopally.*?present Level\s*([\d.]+).*?present Storage\s*([\d.]+)',
        text,
        re.I | re.S
    )

    if a:
        data["Ashwaraopally"] = [
            float(a.group(1)),
            float(a.group(2))
        ]

    # Tapaspally
    t = re.search(
        r'Tapaspally.*?present Level\s*([\d.]+).*?present Storage\s*([\d.]+)',
        text,
        re.I | re.S
    )

    if t:
        data["Tapaspally"] = [
            float(t.group(1)),
            float(t.group(2))
        ]

    # RS Ghanpur
    rsg = re.search(
        r'RS Ghanpur.*?Present Level.*?\+([\d.]+).*?Present Storage\s*:\s*([\d.]+)',
        text,
        re.I | re.S
    )

    if rsg:
        data["R.S. Ghanpur"] = [
            float(rsg.group(1)),
            float(rsg.group(2))
        ]

    # Nawabpet
    n = re.search(
        r'Nawabpet.*?Present Level\s*:\s*\+?([\d.]+).*?\(([\d.]+)',
        text,
        re.I | re.S
    )

    if n:
        data["Nawabpet"] = [
            float(n.group(1)),
            float(n.group(2))
        ]

    # Chitakodur
    c = re.search(
        r'Chitakodur.*?present Level\s*:\s*\+?([\d.]+).*?present Storage\s*:\s*([\d.]+)',
        text,
        re.I | re.S
    )

    if c:
        data["Cheetakodur"] = [
            float(c.group(1)),
            float(c.group(2))
        ]

    # Mylaram
    m = re.search(
        r'Mylaram.*?Present capacity\s*:\s*([\d.]+)TMC.*?\+([\d.]+)',
        text,
        re.I | re.S
    )

    if m:

        level = float(m.group(2))

        storage_mcft = float(m.group(1)) * 1000

        data["Mylaram Balancing Reservoir"] = [
            level,
            storage_mcft
        ]

    return date, data


@app.route("/", methods=["GET", "POST"])
def home():

    table_html = ""
    date = ""
    message = ""
    download_file = ""

    if request.method == "POST":

        action = request.form.get("action")

        message = request.form.get("message", "")

        date, data = parse_messages(message)

        for reservoir, values in data.items():

            table_html += f"""
            <tr>
                <td>{reservoir}</td>
                <td>{values[0]}</td>
                <td>{values[1]}</td>
            </tr>
            """

        if action == "generate":

            wb = load_workbook("master.xlsx")

            sheet = wb["Water levels Major"]

            sheet["A1"] = (
                f"DAILY WATER LEVELS IN RESERVOIRS "
                f"UNDER IRRIGATION CIRCLE, JANGAON\nDATED: {date}"
            )

            for row in range(1, sheet.max_row + 1):

                reservoir = sheet[f'B{row}'].value

                if reservoir is None:
                    continue

                excel_name = str(reservoir).strip()

                if excel_name in data:

                    level = data[excel_name][0]
                    storage = data[excel_name][1]

                    sheet[f'G{row}'] = level
                    sheet[f'H{row}'] = storage

            # Remove unwanted template columns before saving the generated file
            # These are: Inflow In cusecs, Outflow in Cusecs, Rain Fall in mm, Gross capacity in TMC,
            # Ayacut in Acrs, Inflows, Outflows
            sheet.delete_cols(9, 4)
            sheet.delete_cols(11, 3)

            filename = f"SE_IC_JGN_MAJOR_WATER_LEVELS_{date}.xlsx"

            filepath = os.path.join(
                OUTPUT_FOLDER,
                filename
            )

            wb.save(filepath)

            download_file = filename

    return render_template(
        "index.html",
        table=table_html,
        date=date,
        message=message,
        download_file=download_file
    )


@app.route("/download/<filename>")
def download(filename):

    return send_from_directory(
        OUTPUT_FOLDER,
        filename,
        as_attachment=True
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)