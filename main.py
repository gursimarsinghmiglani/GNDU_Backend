import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
from tkinter import Tk, Label, Button, filedialog, messagebox
from tkinter.ttk import Progressbar


def get_official_website(nbfc_name):
    query = f"{nbfc_name} official site"
    url = f"https://www.google.com/search?q={query}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
    }
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")
    try:
        result = soup.find('a')['href']
        return result
    except (AttributeError, TypeError):
        return None


def process_nbfc_file(file_path, progress):
    df = pd.read_excel(file_path)
    df['Official Website'] = None
    for index, row in df.iterrows():
        nbfc_name = row['NBFC Name']
        official_website = get_official_website(nbfc_name)
        df.at[index, 'Official Website'] = official_website
        time.sleep(1)  # To avoid being blocked by Google
        progress['value'] = (index + 1) / len(df) * 100
        root.update_idletasks()
    output_file_path = file_path.lower().replace('.xlsx', '_with_websites.xlsx')
    df.to_excel(output_file_path, index=False)
    messagebox.showinfo("Success", f"Output saved to {output_file_path}")


def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        progress_bar['value'] = 0
        process_nbfc_file(file_path, progress_bar)


root = Tk()
root.title("NBFC Official Website Finder")
Label(root, text="NBFC Official Website Finder", font=("Helvetica", 16)).pack(pady=20)
Button(root, text="Upload Excel File", command=open_file).pack(pady=10)
progress_bar = Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=20)
root.mainloop()
