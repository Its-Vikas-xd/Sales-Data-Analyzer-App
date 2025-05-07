import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Modern color scheme
COLORS = {
    'background': '#2D2D2D',
    'foreground': '#FFFFFF',
    'accent1': '#00B4D8',  # Teal
    'accent2': '#FF6B6B',   # Coral
    'secondary': '#4A4A4A',
    'text': '#E0E0E0',
    'chart_bg': '#1E1E1E'
}

class DataAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sales Data Analyzer")
        self.root.geometry("1000x800")
        self.file_path = None
        self.df = None
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.configure_styles()
        
        # Create main container
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # File selection section
        file_frame = ttk.Frame(self.main_frame)
        file_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(file_frame, text="Select Excel File:", style='Header.TLabel').pack(side=tk.LEFT, padx=5)
        self.file_entry = ttk.Entry(file_frame, width=50, style='Custom.TEntry')
        self.file_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(file_frame, text="Browse", command=self.load_file, style='Accent.TButton').pack(side=tk.LEFT)
        
        # Analysis buttons
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(pady=15)
        
        buttons = [
            ("Data Overview", self.show_data_overview, '#00B4D8'),
            ("Sales Visualizations", self.show_viz_options, '#FF6B6B'),
            ("Exit", self.root.destroy, '#4A4A4A')
        ]
        
        for text, command, color in buttons:
            btn = ttk.Button(button_frame, text=text, command=command, 
                           style=f'Custom.TButton')
            btn.pack(side=tk.LEFT, padx=10, ipadx=10, ipady=5)
        
        # Output console
        console_frame = ttk.LabelFrame(self.main_frame, text="Analysis Output", style='Custom.TLabelframe')
        console_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.output_console = scrolledtext.ScrolledText(console_frame, height=15, wrap=tk.WORD,
                                                      bg=COLORS['secondary'], fg=COLORS['text'],
                                                      insertbackground=COLORS['text'])
        self.output_console.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
    def configure_styles(self):
        self.style.configure('.', background=COLORS['background'], foreground=COLORS['text'])
        self.style.configure('Custom.TEntry', fieldbackground=COLORS['secondary'],
                             foreground=COLORS['text'], bordercolor=COLORS['accent1'],
                             lightcolor=COLORS['accent1'], darkcolor=COLORS['accent1'])
        self.style.map('Custom.TButton',
                       foreground=[('active', COLORS['background']), ('!active', COLORS['text'])],
                       background=[('active', COLORS['accent1']), ('!active', COLORS['secondary'])],
                       bordercolor=[('active', COLORS['accent1'])])
        self.style.configure('Header.TLabel', font=('Helvetica', 12, 'bold'),
                             foreground=COLORS['accent1'])
        self.style.configure('Custom.TLabelframe', background=COLORS['background'],
                             foreground=COLORS['accent1'], bordercolor=COLORS['secondary'])

        plt.style.use('dark_background')
        sns.set_style("darkgrid", {'axes.facecolor': COLORS['chart_bg']})
        plt.rcParams['axes.edgecolor'] = COLORS['text']
        plt.rcParams['text.color'] = COLORS['text']
        plt.rcParams['axes.labelcolor'] = COLORS['text']
        plt.rcParams['xtick.color'] = COLORS['text']
        plt.rcParams['ytick.color'] = COLORS['text']

    def load_file(self):
        filetypes = (("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        filename = filedialog.askopenfilename(title="Open File", filetypes=filetypes)
        if filename:
            try:
                self.df = pd.read_excel(filename)
                self.file_path = filename
                self.file_entry.delete(0, tk.END)
                self.file_entry.insert(0, filename)
                messagebox.showinfo("Success", "File loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")

    def show_error(self, msg):
        messagebox.showerror("Error", msg)

    def show_data_overview(self):
        if self.df is None:
            self.show_error("Please load a file first.")
            return
        overview = self.df.describe(include='all').to_string()
        self.output_console.delete(1.0, tk.END)
        self.output_console.insert(tk.END, overview)

    def show_viz_options(self):
        if self.df is None:
            self.show_error("Please load a file first.")
            return
        viz_window = tk.Toplevel(self.root)
        viz_window.title("Select Visualization")
        viz_window.configure(bg=COLORS['background'])

        options = [
            ("1. Monthly Sales Trend", lambda: self.generate_chart(1)),
            ("2. Top Selling Products", lambda: self.generate_chart(2)),
            ("3. Monthly Sales Box Plot", lambda: self.generate_chart(3)),
            ("4. Individual Product Trends", lambda: self.generate_chart(4)),
        ]

        for label, func in options:
            btn = ttk.Button(viz_window, text=label, command=func, style='Accent.TButton')
            btn.pack(padx=20, pady=10, fill=tk.X)

    def generate_chart(self, chart_choice):
        try:
            months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC']
            self.df[months] = self.df[months].apply(pd.to_numeric, errors='coerce').fillna(0)

            fig = plt.figure(figsize=(10, 6), facecolor=COLORS['chart_bg'])
            ax = fig.add_subplot(111)
            ax.set_facecolor(COLORS['chart_bg'])

            if chart_choice == 1:
                monthly_totals = self.df[months].sum()
                monthly_totals.plot(kind='line', marker='o', color=COLORS['accent1'], ax=ax)
                ax.set_title("Monthly Sales Trend (All Products)", color=COLORS['accent1'])

            elif chart_choice == 2:
                top_products = self.df.nlargest(5, 'Total Sales')
                (top_products.set_index('Electrical Items')[months].sum(axis=1)
                 .sort_values().plot(kind='barh', color=COLORS['accent2'], ax=ax))
                ax.set_title("Top Selling Products (Annual Total)", color=COLORS['accent1'])

            elif chart_choice == 3:
                self.df[months].plot(kind='box', patch_artist=True,
                                     boxprops=dict(facecolor=COLORS['accent1']),
                                     ax=ax)
                ax.set_title("Monthly Sales Distribution", color=COLORS['accent1'])

            elif chart_choice == 4:
                fig.clf()
                top_5 = self.df.nlargest(6, 'Total Sales')
                fig, axes = plt.subplots(3, 2, figsize=(15, 10), facecolor=COLORS['chart_bg'])
                axes = axes.flatten()

                for idx, (_, row) in enumerate(top_5.iterrows()):
                    axes[idx].plot(months, row[months], marker='o', color=COLORS['accent1'])
                    axes[idx].set_title(row['Electrical Items'], fontsize=8, color=COLORS['accent1'])
                    axes[idx].tick_params(axis='x', rotation=45, colors=COLORS['text'])
                    axes[idx].set_facecolor(COLORS['chart_bg'])
                    axes[idx].grid(True, alpha=0.3)

                plt.suptitle("Top Products Monthly Sales Trends", color=COLORS['accent1'])
                plt.tight_layout()

            plt.tight_layout()

            chart_window = tk.Toplevel(self.root)
            chart_window.title("Analysis Result")
            chart_window.configure(bg=COLORS['background'])

            canvas = FigureCanvasTkAgg(fig, master=chart_window)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        except Exception as e:
            self.show_error(f"Chart Error: {str(e)}")

# Run the app
if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalysisApp(root)
    root.mainloop()
