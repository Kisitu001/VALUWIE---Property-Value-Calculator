# Import necessary Python modules
import sys                      # Used for system operations (like safely closing the app on error)
import pandas as pd             # Used for reading and manipulating Excel datasets
import tkinter as tk            # Used for creating desktop graphical user interface windows
from tkinter import ttk, messagebox  # Used for advanced styled widgets (comboboxes, tabs, and popups)

class MailoValuationPortalApp:
    """
    Main application class for the Bulemezi Mailo Calculator & Agents Portal.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("Bulemezi Mailo Calculator & Agents Portal")
        self.root.geometry("560x660")
        self.root.config(bg="#f4f6f9")  # Clean light professional background

        # ==========================================
        # 1. LOAD AND PREPARE DATA FROM EXCEL
        # ==========================================
        try:
            self.excel_file = "Mailo Land Valuation Luwero.xlsx"
            
            # Read both sheets from our workbook
            self.mailo_df = pd.read_excel(self.excel_file, sheet_name="Mailo Land (Luwero MZO)")
            self.agents_df = pd.read_excel(self.excel_file, sheet_name="Agents Directory")
            
            # Clean string formatting across administrative columns to ensure clean matching
            for col in ['District', 'County', 'Subcounty', 'Village / Zone']:
                if col in self.mailo_df.columns:
                    self.mailo_df[col] = self.mailo_df[col].fillna('General').astype(str).str.strip()
            
            # Remove trailing decimal points from block numbers (e.g., converts '1.0' to '1')
            if 'Block No' in self.mailo_df.columns:
                self.mailo_df['Block No'] = self.mailo_df['Block No'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
            
        except Exception as e:
            # If the Excel file is missing or corrupted, show an error box and safely close the app
            messagebox.showerror(
                "Error", 
                f"Could not load Excel dataset:\nMake sure 'Mailo Land Valuation Luwero.xlsx' is in the folder.\n\nDetails: {e}"
            )
            sys.exit(1)

        # Application Header Title Label
        tk.Label(
            root, 
            text="Mailo Valuation Portal — Greater Luwero MZO", 
            font=("Arial", 11, "bold"), 
            fg="#1B365D", 
            bg="#f4f6f9"
        ).pack(pady=8)

        # ==========================================
        # 2. CREATE TABBED INTERFACE (NOTEBOOK)
        # ==========================================
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)

        # Initialize the two main tabs of the application
        self.init_calculator_tab()
        self.init_agents_tab()

    def init_calculator_tab(self):
        """
        Builds the Valuation Portal tab, including the 5-tier cascaded dropdowns,
        input fields for land size and building values, calculation logic, and output report box.
        """
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=" Valuation Portal ")

        # Container frame for input fields
        f = tk.Frame(tab, bg="#f4f6f9")
        f.pack(padx=15, pady=8, fill="x")

        # --- Tier 1: District Selection ---
        tk.Label(f, text="Select District:", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="w", pady=4)
        districts = sorted([d for d in self.mailo_df['District'].unique() if d and d.lower() != 'nan'])
        self.dist_cb = ttk.Combobox(f, values=districts, state="readonly", width=26)
        self.dist_cb.grid(row=0, column=1, pady=4)
        self.dist_cb.bind("<<ComboboxSelected>>", self.on_district_change)

        # --- Tier 2: County Selection ---
        tk.Label(f, text="Select County:", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=1, column=0, sticky="w", pady=4)
        self.county_cb = ttk.Combobox(f, state="readonly", width=26)
        self.county_cb.grid(row=1, column=1, pady=4)
        self.county_cb.bind("<<ComboboxSelected>>", self.on_county_change)

        # --- Tier 3: Subcounty Selection ---
        tk.Label(f, text="Select Subcounty:", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=2, column=0, sticky="w", pady=4)
        self.sub_cb = ttk.Combobox(f, state="readonly", width=26)
        self.sub_cb.grid(row=2, column=1, pady=4)
        self.sub_cb.bind("<<ComboboxSelected>>", self.on_subcounty_change)

        # --- Tier 4: Block Number Selection ---
        tk.Label(f, text="Select Block No:", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=3, column=0, sticky="w", pady=4)
        self.block_cb = ttk.Combobox(f, state="readonly", width=26)
        self.block_cb.grid(row=3, column=1, pady=4)
        self.block_cb.bind("<<ComboboxSelected>>", self.on_block_change)

        # --- Tier 5: Village / Zone Selection ---
        tk.Label(f, text="Select Village / Zone:", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=4, column=0, sticky="w", pady=4)
        self.zone_cb = ttk.Combobox(f, state="readonly", width=26)
        self.zone_cb.grid(row=4, column=1, pady=4)

        # --- Land Size (Acres) Input Entry ---
        tk.Label(f, text="Land Size (Acres):", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=5, column=0, sticky="w", pady=4)
        self.acres_e = tk.Entry(f, width=28)
        self.acres_e.grid(row=5, column=1, pady=4)
        self.acres_e.insert(0, "1.0")

        # --- Building / Improvements Value Input Entry ---
        tk.Label(f, text="Building Value (UGX):", bg="#f4f6f9", font=("Arial", 9, "bold")).grid(row=6, column=0, sticky="w", pady=4)
        self.bldg_e = tk.Entry(f, width=28)
        self.bldg_e.grid(row=6, column=1, pady=4)
        self.bldg_e.insert(0, "0")

        if districts:
            self.dist_cb.current(0)
            self.on_district_change(None)

        # Calculate Action Button
        tk.Button(
            tab, 
            text="Calculate Fees & Stamp Duty", 
            command=self.calculate, 
            bg="#1B365D", 
            fg="white", 
            font=("Arial", 9, "bold"), 
            width=36, 
            pady=4
        ).pack(pady=6)

        # Output Report Text Box (Read-only)
        self.res = tk.Text(tab, height=10, width=58, font=("Courier", 9), bg="#ffffff")
        self.res.pack(padx=15, pady=2)
        self.res.insert(tk.END, "Select location hierarchy and click calculate...")
        self.res.config(state="disabled")

    # ==========================================
    # 3. CASCADED DROPDOWN UPDATE LOGIC
    # ==========================================
    def on_district_change(self, event):
        dist = self.dist_cb.get()
        counties = sorted(self.mailo_df[self.mailo_df['District'] == dist]['County'].unique().tolist())
        self.county_cb['values'] = counties
        if counties:
            self.county_cb.current(0)
            self.on_county_change(None)

    def on_county_change(self, event):
        dist, county = self.dist_cb.get(), self.county_cb.get()
        subcounties = sorted(self.mailo_df[(self.mailo_df['District'] == dist) & (self.mailo_df['County'] == county)]['Subcounty'].unique().tolist())
        self.sub_cb['values'] = subcounties
        if subcounties:
            self.sub_cb.current(0)
            self.on_subcounty_change(None)

    def on_subcounty_change(self, event):
        dist, county, sub = self.dist_cb.get(), self.county_cb.get(), self.sub_cb.get()
        f_df = self.mailo_df[(self.mailo_df['District'] == dist) & (self.mailo_df['County'] == county) & (self.mailo_df['Subcounty'] == sub)]
        blocks = sorted(f_df['Block No'].unique().tolist(), key=lambda x: int(x) if x.isdigit() else 9999)
        self.block_cb['values'] = blocks
        if blocks:
            self.block_cb.current(0)
            self.on_block_change(None)

    def on_block_change(self, event):
        dist, county, sub, block = self.dist_cb.get(), self.county_cb.get(), self.sub_cb.get(), self.block_cb.get()
        f_df = self.mailo_df[
            (self.mailo_df['District'] == dist) & 
            (self.mailo_df['County'] == county) & 
            (self.mailo_df['Subcounty'] == sub) & 
            (self.mailo_df['Block No'] == block)
        ]
        villages = sorted(f_df['Village / Zone'].unique().tolist())
        self.zone_cb['values'] = villages
        if villages:
            self.zone_cb.current(0)

    # ==========================================
    # 4. VALUATION & TAX CALCULATION LOGIC
    # ==========================================
    def calculate(self):
        try:
            dist = self.dist_cb.get()
            county = self.county_cb.get()
            sub = self.sub_cb.get()
            block = self.block_cb.get()
            zone = self.zone_cb.get()
            
            acres = float(self.acres_e.get())
            bldg = float(self.bldg_e.get())
            
            match = self.mailo_df[
                (self.mailo_df['District'] == dist) & 
                (self.mailo_df['County'] == county) & 
                (self.mailo_df['Subcounty'] == sub) & 
                (self.mailo_df['Block No'] == block) & 
                (self.mailo_df['Village / Zone'] == zone)
            ]['Fair Market Value (UGX)'].values
            
            rate = match[0] if len(match) > 0 and pd.notna(match[0]) else 10000000

            land_val = rate * acres
            total_prop = land_val + bldg
            stamp_duty = total_prop * 0.015  # 1.5% Stamp Duty
            reg_fee = 20000                   # Registration Fee
            search_fee = 20000                # Official Search Fee
            total_cost = stamp_duty + reg_fee + search_fee

            text = (
                f"==================================================\n"
                f" MAILO VALUATION REPORT ({dist.upper()} - {county.upper()})\n"
                f"==================================================\n"
                f"• Subcounty      : {sub}\n"
                f"• Block Number   : Block {block}\n"
                f"• Village / Zone : {zone}\n"
                f"• Size (Acres)   : {acres} acre(s)\n"
                f"• Bare Land Val  : UGX {land_val:,.2f}\n"
                f"• Building Val   : UGX {bldg:,.2f}\n"
                f"--------------------------------------------------\n"
                f"• Total Property : UGX {total_prop:,.2f}\n\n"
                f"=== STATUTORY TAXES & FEES ===\n"
                f"• Stamp Duty(1.5%): UGX {stamp_duty:,.2f}\n"
                f"• Title Reg. Fee : UGX {reg_fee:,.2f}\n"
                f"• Official Search: UGX {search_fee:,.2f}\n"
                f"--------------------------------------------------\n"
                f"• TOTAL ACQUISITION COST: UGX {total_cost:,.2f}\n"
                f"==================================================\n"
            )
            self.res.config(state="normal")
            self.res.delete("1.0", tk.END)
            self.res.insert(tk.END, text)
            self.res.config(state="disabled")
            
        except Exception as err:
            messagebox.showerror("Input Error", f"Please enter valid numeric values:\n{err}")

    # ==========================================
    # 5. FIELD AGENTS DIRECTORY TAB
    # ==========================================
    def init_agents_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text=" Field Agents Directory (15) ")

        f = tk.Frame(tab, bg="#f4f6f9", padx=10, pady=10)
        f.pack(fill="both", expand=True)

        tk.Label(f, text="Verified Luwero MZO Field Agents & Fees:", font=("Arial", 9, "bold"), fg="#1B365D", bg="#f4f6f9").pack(anchor="w", pady=(0, 5))

        table_frame = tk.Frame(f)
        table_frame.pack(fill="both", expand=True)

        scroll_y = ttk.Scrollbar(table_frame, orient="vertical")
        scroll_x = ttk.Scrollbar(table_frame, orient="horizontal")

        cols = list(self.agents_df.columns)
        tree = ttk.Treeview(table_frame, columns=cols, show="headings", yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.config(command=tree.yview)
        scroll_x.config(command=tree.xview)

        scroll_y.pack(side="right", fill="y")
        scroll_x.pack(side="bottom", fill="x")
        tree.pack(side="left", fill="both", expand=True)

        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=140, anchor="w")

        for _, row in self.agents_df.iterrows():
            vals = [f"{v:,.0f}" if isinstance(v, (int, float)) and "Fee" in c else str(v) for c, v in zip(cols, row)]
            tree.insert("", "end", values=vals)

if __name__ == "__main__":
    root = tk.Tk()
    app = MailoValuationPortalApp(root)
    root.mainloop()