import pandas as pd
import FreeSimpleGUI as sg


layout = [
    [sg.Text("Filename")],
    [sg.Input(), sg.FileBrowse()],
    [sg.OK(), sg.Cancel()],
]
window = sg.Window("File Splitter: Ver 2.0", layout)

event, file = window.Read()


def reader(file):
    dt = pd.read_excel(file)
    dt = dt.drop(dt.tail(1).index)

    columns_to_keep = [
        "Full Name",
        "PID",
        "Mjr/Min/Cert",
        "Mjr/Min/Cert Code",
        "Dept/School",
        "Mailing Address 1",
        "Mailing Address City",
        "Mailing Address State",
        "Mailing Address Zip",
        "Advisor(s)",
    ]

    dt = dt[columns_to_keep]
    dt.columns = [
        "Name",
        "PID",
        "Major",
        "Major_Code",
        "Dept",
        "Address",
        "City",
        "State",
        "Zip",
        "Advisor(s)",
    ]

    phd_codes = dt.Major_Code.str.contains("^PH|^ED")
    sports_codes = dt.Dept.str.contains(
        r"Recreation and Sport Pedagogy|Rec, Sport Ped & Cons Sci", regex=True
    )

    phd = dt[phd_codes].sort_values(by=["Major_Code", "Name"])
    sports = dt[sports_codes].sort_values(by=["Major_Code", "Name"])
    rest = dt[~phd_codes & ~sports_codes].sort_values(by=["Major_Code", "Name"])

    with pd.ExcelWriter("Output.xlsx", engine="xlsxwriter") as writer:
        dt.to_excel(writer, sheet_name="All", index=False, freeze_panes=(1, 0))
        phd.to_excel(writer, sheet_name="PHD", index=False, freeze_panes=(1, 0))
        rest.to_excel(writer, sheet_name="Masters", index=False, freeze_panes=(1, 0))
        sports.to_excel(
            writer, sheet_name="Recreation", index=False, freeze_panes=(1, 0)
        )

        # adding colors to the tabs
        tabs = ["All", "PHD", "Masters", "Recreation"]
        colors = ["red", "green", "orange", "blue"]

        for sheet, color in zip(tabs, colors):
            sh = writer.sheets[sheet]
            sh.set_tab_color(color)

        dfs = [dt, phd, rest, sports]

        for sheet, df in zip(tabs, dfs):
            sh = writer.sheets[sheet]
            for idx, col in enumerate(df):
                max_len = max(df[col].apply(lambda x: len(str(x))).max(), len(col))
                sh.set_column(idx, idx, max_len)


fname = file[0]

if __name__ == "__main__":
    if fname:
        try:
            reader(fname)
            sg.popup(
                "Done!\nThe file Output.xlsx is saved in the same location.",
                title="Completed",
            )
        except KeyError as k:
            sg.popup_error(
                f"Error! Check the content of the input Excel file\nColumn(s) may need to be renamed or they do not exist\n{k}",
            )
        except Exception as e:
            sg.popup_error(f"An error occurred:\n{e}")
