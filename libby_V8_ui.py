import customtkinter as ctk
import json
import os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "libby_config.json")

DEFAULT_CONFIG = {
    "assistant_name": "Libby",
    "knowledge_folder": SCRIPT_DIR,
    "theme": "dark",
    "company_name": "My Company",
    "active_modules": ["knowledge", "bi"]
}

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return DEFAULT_CONFIG.copy()

def save_config(config):
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)

class LibbyApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.config_data = load_config()
        self.assistant_name = self.config_data.get("assistant_name", "Libby")

        ctk.set_appearance_mode(self.config_data.get("theme", "dark"))
        ctk.set_default_color_theme("dark-blue")

        self.title(f"{self.assistant_name} — AI Knowledge Assistant")
        self.geometry("1100x700")
        self.minsize(800, 500)

        self.ROSE_GOLD = "#c9956c"
        self.build_layout()

    def build_layout(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.build_sidebar()
        self.build_main()

    def build_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)

        name_label = ctk.CTkLabel(
            self.sidebar,
            text=self.assistant_name.upper(),
            font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
            text_color=self.ROSE_GOLD
        )
        name_label.pack(pady=(20, 2), padx=16, anchor="w")

        sub_label = ctk.CTkLabel(
            self.sidebar,
            text="powered by local AI",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color="gray"
        )
        sub_label.pack(padx=16, anchor="w")

        ctk.CTkFrame(self.sidebar, height=1, fg_color="gray30").pack(
            fill="x", padx=12, pady=12
        )

        self.kb_btn = ctk.CTkButton(
            self.sidebar, text="📚  Knowledge",
            anchor="w", fg_color="transparent",
            text_color=self.ROSE_GOLD,
            hover_color="gray25",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            command=lambda: self.switch_tab(0)
        )
        self.kb_btn.pack(fill="x", padx=8, pady=2)

        self.bi_btn = ctk.CTkButton(
            self.sidebar, text="📊  Enterprise BI",
            anchor="w", fg_color="transparent",
            text_color="gray",
            hover_color="gray25",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            command=lambda: self.switch_tab(1)
        )
        self.bi_btn.pack(fill="x", padx=8, pady=2)

        settings_btn = ctk.CTkButton(
            self.sidebar, text="⚙  Settings",
            anchor="w", fg_color="transparent",
            text_color="gray",
            hover_color="gray25",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            command=self.open_settings
        )
        settings_btn.pack(fill="x", padx=8, pady=2, side="bottom")

        ctk.CTkFrame(self.sidebar, height=1, fg_color="gray30").pack(
            fill="x", padx=12, pady=8, side="bottom"
        )

    def build_main(self):
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew")
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        self.tabview = ctk.CTkTabview(
            self.main_frame,
            segmented_button_selected_color=self.ROSE_GOLD,
            segmented_button_selected_hover_color=self.ROSE_GOLD
        )
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        self.main_frame.grid_rowconfigure(0, weight=1)

        self.tab_kb = self.tabview.add("📚 Knowledge Assistant")
        self.tab_bi = self.tabview.add("📊 Enterprise BI")

        self.build_chat_tab(self.tab_kb, "knowledge")
        self.build_chat_tab(self.tab_bi, "bi")

    def build_chat_tab(self, tab, tab_id):
        tab.grid_rowconfigure(0, weight=1)
        tab.grid_columnconfigure(0, weight=1)

        chat_frame = ctk.CTkScrollableFrame(tab, label_text="")
        chat_frame.grid(row=0, column=0, sticky="nsew", padx=8, pady=(8, 4))

        if tab_id == "knowledge":
            self.kb_chat_frame = chat_frame
        else:
            self.bi_chat_frame = chat_frame

        input_frame = ctk.CTkFrame(tab, fg_color="transparent")
        input_frame.grid(row=1, column=0, sticky="ew", padx=8, pady=(4, 8))
        input_frame.grid_columnconfigure(0, weight=1)

        entry = ctk.CTkEntry(
            input_frame,
            placeholder_text=f"Ask {self.assistant_name} something...",
            font=ctk.CTkFont(family="Segoe UI", size=13),
            height=40,
            corner_radius=8
        )
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        send_btn = ctk.CTkButton(
            input_frame,
            text="Send",
            width=80,
            height=40,
            corner_radius=8,
            fg_color=self.ROSE_GOLD,
            hover_color="#a87850",
            text_color="#1e1e1e",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            command=lambda: self.send_message(entry, tab_id)
        )
        send_btn.grid(row=0, column=1)

        entry.bind("<Return>", lambda e: self.send_message(entry, tab_id))

        if tab_id == "knowledge":
            self.kb_entry = entry
        else:
            self.bi_entry = entry

        self.add_bubble(
            chat_frame,
            f"Hello! I'm {self.assistant_name}. Ask me anything about your documents.",
            is_user=False,
            show_feedback=False
        )

    def add_bubble(self, chat_frame, text, is_user=False, source=None, show_feedback=True):
        bubble_frame = ctk.CTkFrame(
            chat_frame,
            fg_color="#3a2e2e" if is_user else "gray20",
            corner_radius=12
        )
        bubble_frame.pack(
            anchor="e" if is_user else "w",
            pady=4, padx=8,
            fill="none"
        )

        msg_label = ctk.CTkLabel(
            bubble_frame,
            text=text,
            font=ctk.CTkFont(family="Segoe UI", size=13),
            wraplength=500,
            justify="left",
            text_color="#dddddd"
        )
        msg_label.pack(padx=12, pady=(10, 4), anchor="w")

        if source:
            src_label = ctk.CTkLabel(
                bubble_frame,
                text=f"Source: {source}",
                font=ctk.CTkFont(family="Segoe UI", size=10),
                text_color=self.ROSE_GOLD
            )
            src_label.pack(padx=12, pady=(0, 4), anchor="w")

        if not is_user and show_feedback:
            fb_frame = ctk.CTkFrame(bubble_frame, fg_color="transparent")
            fb_frame.pack(padx=12, pady=(0, 8), anchor="w")

            ctk.CTkButton(
                fb_frame, text="👍", width=36, height=24,
                fg_color="transparent",
                border_width=1,
                border_color="gray40",
                text_color="gray",
                hover_color="gray25",
                font=ctk.CTkFont(size=11),
                command=lambda: self.log_feedback(text, "positive")
            ).pack(side="left", padx=(0, 4))

            ctk.CTkButton(
                fb_frame, text="👎", width=36, height=24,
                fg_color="transparent",
                border_width=1,
                border_color="gray40",
                text_color="gray",
                hover_color="gray25",
                font=ctk.CTkFont(size=11),
                command=lambda: self.log_feedback(text, "negative")
            ).pack(side="left")

    def send_message(self, entry, tab_id):
        text = entry.get().strip()
        if not text:
            return
        entry.delete(0, "end")

        chat_frame = self.kb_chat_frame if tab_id == "knowledge" else self.bi_chat_frame
        self.add_bubble(chat_frame, text, is_user=True, show_feedback=False)

        response = f"[Libby engine not connected yet — RAG coming in next build]"
        self.add_bubble(chat_frame, response, is_user=False, source="placeholder")

    def log_feedback(self, answer_text, rating):
        import datetime
        log_file = os.path.join(SCRIPT_DIR, "feedback_log.json")
        entry = {
            "timestamp": datetime.datetime.now().isoformat(),
            "answer": answer_text,
            "rating": rating
        }
        log = []
        if os.path.exists(log_file):
            with open(log_file, "r") as f:
                log = json.load(f)
        log.append(entry)
        with open(log_file, "w") as f:
            json.dump(log, f, indent=2)
        print(f"Feedback logged: {rating}")

    def switch_tab(self, index):
        tabs = ["📚 Knowledge Assistant", "📊 Enterprise BI"]
        self.tabview.set(tabs[index])

    def open_settings(self):
        settings_win = ctk.CTkToplevel(self)
        settings_win.title("Settings")
        settings_win.geometry("420x320")
        settings_win.grab_set()

        ctk.CTkLabel(
            settings_win,
            text="Settings",
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold")
        ).pack(pady=(20, 16), padx=20, anchor="w")

        ctk.CTkLabel(settings_win, text="Assistant Name").pack(anchor="w", padx=20)
        name_entry = ctk.CTkEntry(settings_win, width=300)
        name_entry.insert(0, self.config_data.get("assistant_name", "Libby"))
        name_entry.pack(padx=20, pady=(4, 12))

        ctk.CTkLabel(settings_win, text="Knowledge Folder").pack(anchor="w", padx=20)
        folder_entry = ctk.CTkEntry(settings_win, width=300)
        folder_entry.insert(0, self.config_data.get("knowledge_folder", SCRIPT_DIR))
        folder_entry.pack(padx=20, pady=(4, 12))

        ctk.CTkLabel(settings_win, text="Theme").pack(anchor="w", padx=20)
        theme_menu = ctk.CTkOptionMenu(settings_win, values=["dark", "light", "system"])
        theme_menu.set(self.config_data.get("theme", "dark"))
        theme_menu.pack(padx=20, pady=(4, 20), anchor="w")

        def save_settings():
            self.config_data["assistant_name"] = name_entry.get()
            self.config_data["knowledge_folder"] = folder_entry.get()
            self.config_data["theme"] = theme_menu.get()
            save_config(self.config_data)
            ctk.set_appearance_mode(theme_menu.get())
            settings_win.destroy()

        ctk.CTkButton(
            settings_win, text="Save",
            fg_color=self.ROSE_GOLD,
            hover_color="#a87850",
            text_color="#1e1e1e",
            command=save_settings
        ).pack(pady=8)

if __name__ == "__main__":
    app = LibbyApp()
    app.mainloop()



