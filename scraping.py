from googlesearch import search
import pandas as pd
import tkinter as tk
from tkinter import simpledialog, messagebox, filedialog

def search_google(query, num_results=10):
    try:
        result_urls = [url for url in search(query, num_results=num_results)]
        return result_urls
    except Exception as e:
        print(f"An error occurred during Google search: {e}")
        return []

def get_search_parameters():
    root = tk.Tk()
    root.withdraw()

    query = simpledialog.askstring("Google Search", "Enter the search query:")
    if query is None:
        return None, None  # Exit if user cancels

    num_results = simpledialog.askinteger("Google Search", "Enter the number of results to fetch:", minvalue=1)
    return query, num_results

def main():
    query, num_results = get_search_parameters()
    
    if query is None or num_results is None:
        messagebox.showinfo("Google Search", "Search parameters not provided. Exiting.")
        return
    
    result_urls = search_google(query, num_results)
    
    if not result_urls:
        messagebox.showinfo("Google Search", "No results found.")
        return
    
    df = pd.DataFrame(result_urls, columns=["URLs"])
    
    excel_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    
    if excel_file:
        try:
            df.to_excel(excel_file, engine='xlsxwriter', index=False)
            messagebox.showinfo("Google Search", f"Results saved in {excel_file}")
        except Exception as e:
            messagebox.showinfo("Google Search", f"An error occurred while saving results: {e}")
    else:
        messagebox.showinfo("Google Search", "No file selected. Results not saved.")

if __name__ == "__main__":
    main()
