import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import requests
import time
import os
import webbrowser
from datetime import datetime
import csv

########################################################################
# GLOBAL CONSTANTS
########################################################################
OAUTH_URL           = "https://api.connect.withsecure.com/as/token.oauth2"
ORGANIZATIONS_URL   = "https://api.connect.withsecure.com/organizations/v1/organizations"
DEVICES_URL         = "https://api.connect.withsecure.com/devices/v1/devices"
OPERATIONS_URL      = "https://api.connect.withsecure.com/devices/v1/operations"
PATCH_DEVICES_URL   = "https://api.connect.withsecure.com/devices/v1/devices"  # block/inactivate
DELETE_DEVICES_URL  = "https://api.connect.withsecure.com/devices/v1/devices"  # delete
MISSING_UPDATES_URL = "https://api.connect.withsecure.com/software-updates/v1/missing-updates"
OAUTH_SCOPE         = "connect.api.read connect.api.write"
USER_AGENT          = "WithSecureBatchManagement/1.0"

########################################################################
# PREDEFINED OPERATIONS
# (the user can't edit these from the combo box, it's read-only)
########################################################################
OPERATION_DESCRIPTIONS = {
    "scanForMalware": (
        "Initiates an anti-malware scan on each selected device.\n"
        "Requires the endpoint to have an EPP (Computer Protection) license."
    ),
    "showMessage": (
        "Sends a custom pop-up message to each device's user.\n"
        "Requires EPP licensing and an active device."
    ),
    "isolateFromNetwork": (
        "Isolates selected devices from the network.\n"
        "Requires an active firewall (EPP)."
    ),
    "releaseFromNetworkIsolation": (
        "Reverses an earlier isolation, restoring normal network access.\n"
        "Applies only if the device was previously isolated."
    ),
    "assignProfile": (
        "Assigns a specified WithSecure profile to the devices.\n"
        "Prompts for a numeric profileId. The device must be active."
    ),
    "turnOnFeature": (
        "Enables 'debugLogging' for 5–1440 minutes on selected devices.\n"
        "Used for troubleshooting. Only for EPP-licensed endpoints.\n"
        "Automatically disables after the specified time or on reboot."
    ),
    "checkMissingUpdates": (
        "Checks for missing software updates on each device.\n"
        "If found, you can export them to a CSV (includes app names)."
    ),
    "inventory": (
        "Exports a CSV containing hardware info about the selected devices.\n"
        "Groups them by organization name, then device name."
    )
}

########################################################################
# CLASS: ToolTip
########################################################################
class ToolTip:
    """
    A small balloon tooltip near the mouse pointer on hover.
    """
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwin = None

        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tipwin or not self.text:
            return
        x = self.widget.winfo_pointerx() + 20
        y = self.widget.winfo_pointery() + 10
        self.tipwin = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")

        label = ttk.Label(tw, text=self.text, background="lightyellow", borderwidth=1, relief="solid")
        label.pack(ipadx=5, ipady=2)

    def hide_tip(self, event=None):
        if self.tipwin:
            self.tipwin.destroy()
        self.tipwin = None


########################################################################
# SCROLLABLE FRAME
########################################################################
def create_scrollable_frame(parent):
    """
    Creates a canvas+scrollbar for a list of checkboxes or items.
    Returns (container, inner_frame, canvas, scrollbar).
    """
    container = ttk.Frame(parent)

    canvas = tk.Canvas(container, highlightthickness=0)
    vscroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
    canvas.configure(yscrollcommand=vscroll.set)

    inner_frame = ttk.Frame(canvas)
    canvas.create_window((0, 0), window=inner_frame, anchor="nw")

    def _on_frame_configure(event):
        canvas.config(scrollregion=canvas.bbox("all"))

    inner_frame.bind("<Configure>", _on_frame_configure)

    # Only scroll if mouse is hovered over
    def _mousewheel_scroll(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _enter(event):
        canvas.bind_all("<MouseWheel>", _mousewheel_scroll)

    def _leave(event):
        canvas.unbind_all("<MouseWheel>")

    canvas.bind("<Enter>", _enter)
    canvas.bind("<Leave>", _leave)

    canvas.pack(side="left", fill="both", expand=True)
    vscroll.pack(side="right", fill="y")

    return container, inner_frame, canvas, vscroll


########################################################################
# MAIN APPLICATION
########################################################################
class WithSecureBatchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("WithSecure Batch Management")
        self.resizable(True, True)

        # SUBTITLE: after .title()
        # we'll place a label for the subtitle in the UI

        # State
        self.client_id_var     = tk.StringVar()
        self.client_secret_var = tk.StringVar()
        self.access_token      = None

        # Orgs
        self.org_vars = []
        self.org_scroll_frame = None
        self.org_inner_frame  = None
        self.org_no_login_label = None

        # Devices
        self.device_vars      = []
        self.dev_scroll_frame = None
        self.dev_inner_frame  = None
        self.dev_no_login_label = None

        # Operation
        self.operation_var    = tk.StringVar(value="scanForMalware")
        self.message_var      = tk.StringVar(value="Greetings from Eykthirnyr !")

        # Logging
        self.log_folder = None

        # list of widgets to disable pre-login
        self.widget_list_to_disable = []

        self._create_widgets()
        self._disable_post_init()

    def _create_widgets(self):
        # Main Title
        lbl_title = ttk.Label(self, text="WithSecure Batch Management", font=("Arial", 16, "bold"))
        lbl_title.pack(pady=5)

        # Sub Title
        lbl_subtitle = ttk.Label(
            self,
            text="Multi-Organization, Multi-Device Bulk Tasks",
            font=("Arial", 10, "italic")
        )
        lbl_subtitle.pack(pady=2)

        # 1) OAuth
        frame_login = ttk.LabelFrame(self, text="1) OAuth Login")
        frame_login.pack(fill="x", padx=5, pady=5)

        ttk.Label(frame_login, text="Client ID:").grid(row=0, column=0, sticky="e", padx=5, pady=2)
        ent_cid = ttk.Entry(frame_login, textvariable=self.client_id_var, width=40)
        ent_cid.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(frame_login, text="Client Secret:").grid(row=1, column=0, sticky="e", padx=5, pady=2)
        ent_sec = ttk.Entry(frame_login, textvariable=self.client_secret_var, width=40, show="*")
        ent_sec.grid(row=1, column=1, padx=5, pady=2)

        # Manage API Keys on the right
        btn_api_keys = ttk.Button(frame_login, text="Manage API Keys", command=self.open_api_keys_page)
        btn_api_keys.grid(row=0, column=2, rowspan=2, sticky="ns", padx=5)

        # Login
        btn_login = ttk.Button(frame_login, text="Login", command=self.handle_login)
        btn_login.grid(row=2, column=0, columnspan=3, sticky="ew", padx=5, pady=5)

        # 2) Orgs
        frame_org = ttk.LabelFrame(self, text="2) Select Organizations")
        frame_org.pack(fill="both", expand=True, padx=5, pady=5)

        self.org_no_login_label = ttk.Label(frame_org, text="Please log in first.")
        self.org_no_login_label.pack()

        org_cont, org_inner, _, _ = create_scrollable_frame(frame_org)
        org_cont.pack(fill="both", expand=True, padx=5, pady=5)
        self.org_scroll_frame = org_cont
        self.org_inner_frame  = org_inner

        org_btn_frame = ttk.Frame(frame_org)
        org_btn_frame.pack(fill="x", padx=5, pady=5)

        btn_org_sel_all = ttk.Button(org_btn_frame, text="Select All Orgs", command=self.select_all_orgs)
        btn_org_sel_all.pack(side="left", padx=5)

        btn_org_des_all = ttk.Button(org_btn_frame, text="Deselect All Orgs", command=self.clear_all_orgs)
        btn_org_des_all.pack(side="left", padx=5)

        btn_org_fetch = ttk.Button(org_btn_frame, text="Fetch Devices", command=self.fetch_devices_for_selected_orgs)
        btn_org_fetch.pack(side="left", padx=5)

        self.widget_list_to_disable.extend([
            org_cont, btn_org_sel_all, btn_org_des_all, btn_org_fetch
        ])

        # 3) Devices
        frame_dev = ttk.LabelFrame(self, text="3) Select Devices")
        frame_dev.pack(fill="both", expand=True, padx=5, pady=5)

        self.dev_no_login_label = ttk.Label(frame_dev, text="Please log in and fetch organizations first.")
        self.dev_no_login_label.pack()

        dev_cont, dev_inner, _, _ = create_scrollable_frame(frame_dev)
        dev_cont.pack(fill="both", expand=True, padx=5, pady=5)
        self.dev_scroll_frame = dev_cont
        self.dev_inner_frame  = dev_inner

        dev_btn_frame = ttk.Frame(frame_dev)
        dev_btn_frame.pack(fill="x", padx=5, pady=5)

        btn_dev_sel_all = ttk.Button(dev_btn_frame, text="Select All Devices", command=self.select_all_devices)
        btn_dev_sel_all.pack(side="left", padx=5)

        btn_dev_des_all = ttk.Button(dev_btn_frame, text="Deselect All Devices", command=self.clear_all_devices)
        btn_dev_des_all.pack(side="left", padx=5)

        self.widget_list_to_disable.extend([
            dev_cont, btn_dev_sel_all, btn_dev_des_all
        ])

        # 4) Operations
        frame_ops = ttk.LabelFrame(self, text="4) Operations on Selected Devices")
        frame_ops.pack(fill="x", padx=5, pady=5)

        lbl_op = ttk.Label(frame_ops, text="Operation:")
        lbl_op.pack(side="left", padx=5, pady=5)
        self.widget_list_to_disable.append(lbl_op)

        # Predefined list, user cannot type new
        ops_keys = list(OPERATION_DESCRIPTIONS.keys())

        self.op_combo = ttk.Combobox(frame_ops, state="readonly", values=ops_keys, textvariable=self.operation_var)
        self.op_combo.pack(side="left", padx=5, pady=5)
        self.op_combo.current(0)
        self.widget_list_to_disable.append(self.op_combo)

        btn_ops_help = ttk.Button(frame_ops, text="?", width=3, command=self.show_operation_description)
        btn_ops_help.pack(side="left", padx=3, pady=5)
        self.widget_list_to_disable.append(btn_ops_help)

        lbl_msg = ttk.Label(frame_ops, text="Message:")
        lbl_msg.pack(side="left", padx=5, pady=5)
        self.widget_list_to_disable.append(lbl_msg)

        msg_entry = ttk.Entry(frame_ops, textvariable=self.message_var, width=30)
        msg_entry.pack(side="left", padx=5, pady=5)
        self.widget_list_to_disable.append(msg_entry)

        btn_trigger = ttk.Button(frame_ops, text="Trigger Operation", command=self.trigger_operation_on_selected)
        btn_trigger.pack(side="left", padx=5, pady=5)
        self.widget_list_to_disable.append(btn_trigger)

        # 5) Device State & Deletion
        frame_state = ttk.LabelFrame(self, text="Device State & Deletion")
        frame_state.pack(fill="x", padx=5, pady=5)

        btn_block = ttk.Button(frame_state, text="Block", command=lambda: self.update_device_state("blocked"))
        btn_block.pack(side="left", padx=5, pady=5)
        ToolTip(btn_block,
            "Block device(s): Freed seat, no data to WithSecure, not visible in portal.\n"
            "Requires EPP license, device must be active."
        )
        self.widget_list_to_disable.append(btn_block)

        btn_inact = ttk.Button(frame_state, text="Inactivate", command=lambda: self.update_device_state("inactive"))
        btn_inact.pack(side="left", padx=5, pady=5)
        ToolTip(btn_inact,
            "Inactivate device(s): Freed seat, not visible in portal.\n"
            "Device re-activates automatically if it reconnects & a seat is free.\n"
            "Requires EPP license, device must be active or blocked."
        )
        self.widget_list_to_disable.append(btn_inact)

        btn_del = ttk.Button(frame_state, text="Delete", command=self.delete_devices)
        btn_del.pack(side="left", padx=5, pady=5)
        ToolTip(btn_del,
            "Delete device(s): Freed seat, permanently removed from the portal.\n"
            "If reinstalled, the device gets a new UUID.\n"
            "Requires EPP license, device must be active."
        )
        self.widget_list_to_disable.append(btn_del)

        # Log Settings
        frame_log = ttk.LabelFrame(self, text="Log Settings")
        frame_log.pack(fill="x", padx=5, pady=5)

        btn_log = ttk.Button(frame_log, text="Choose Log Folder", command=self.choose_log_folder)
        btn_log.pack(side="left", padx=5, pady=5)
        # not disabled

        # About
        frame_about = ttk.LabelFrame(self, text="About")
        frame_about.pack(fill="x", padx=5, pady=5)

        about_frame = ttk.Frame(frame_about)
        about_frame.pack(fill="both", expand=True)

        lbl_about = ttk.Label(
            about_frame,
            text=(
                "Made by Clément GHANEME\n"
                "Use at your own risk.\n"
                "Author not liable for damage or data loss."
            ),
            foreground="gray",
            justify="center"
        )
        lbl_about.pack(pady=5)

        btn_web = ttk.Button(about_frame, text="Visit My Website", command=self.open_website)
        btn_web.pack(pady=5)

    def _disable_post_init(self):
        self.set_widgets_state("disabled")

    def set_widgets_state(self, new_state):
        for w in self.widget_list_to_disable:
            try:
                w.config(state=new_state)
            except:
                # might be a container
                if hasattr(w, "winfo_children"):
                    for c in w.winfo_children():
                        try:
                            c.config(state=new_state)
                        except:
                            pass

    ####################################################################
    # HELPER
    ####################################################################
    def open_api_keys_page(self):
        webbrowser.open("https://elements.withsecure.com/apps/ccr/api_keys")

    def open_website(self):
        webbrowser.open("https://clement.business/")

    def show_operation_description(self):
        op = self.operation_var.get()
        desc = OPERATION_DESCRIPTIONS.get(op, "No description available.")
        messagebox.showinfo(f"Operation: {op}", desc)

    def choose_log_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.log_folder = folder
            messagebox.showinfo("Log Folder", f"Logs will be stored in: {folder}")

    def append_log(self, step, content):
        if not self.log_folder:
            return
        fname = f"{step}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        path  = os.path.join(self.log_folder, fname)
        try:
            with open(path, "a", encoding="utf-8") as f:
                f.write(content + "\n")
        except:
            pass

    ####################################################################
    # 1) LOGIN
    ####################################################################
    def handle_login(self):
        cid = self.client_id_var.get().strip()
        sec = self.client_secret_var.get().strip()
        if not cid or not sec:
            messagebox.showerror("Error", "Client ID and Client Secret required.")
            return
        try:
            auth = (cid, sec)
            data = {"grant_type":"client_credentials", "scope":OAUTH_SCOPE}
            headers = {"User-Agent":USER_AGENT}

            resp = requests.post(OAUTH_URL, auth=auth, data=data, headers=headers)
            self.append_log("login", f"Status={resp.status_code}\nBody={resp.text}")
            resp.raise_for_status()

            j = resp.json()
            self.access_token = j["access_token"]

            messagebox.showinfo("Success", "Login successful.")
            self.set_widgets_state("normal")

            if self.org_no_login_label:
                self.org_no_login_label.destroy()
                self.org_no_login_label = None
            if self.dev_no_login_label:
                self.dev_no_login_label.destroy()
                self.dev_no_login_label = None

            self.fetch_organizations()
        except Exception as e:
            self.append_log("login", f"Error: {e}")
            messagebox.showerror("Error", f"Login failed: {e}")

    ####################################################################
    # 2) FETCH ORGS
    ####################################################################
    def fetch_organizations(self):
        for c in self.org_inner_frame.winfo_children():
            c.destroy()
        self.org_vars.clear()

        if not self.access_token:
            return
        headers = {"User-Agent":USER_AGENT, "Authorization":f"Bearer {self.access_token}"}
        try:
            r = requests.get(ORGANIZATIONS_URL, headers=headers)
            self.append_log("orgs", f"GET => {r.status_code}\n{r.text}")
            r.raise_for_status()

            data = r.json()
            items= data.get("items", [])
            for org in items:
                var = tk.BooleanVar(value=False)
                txt = org["name"]
                cb  = ttk.Checkbutton(self.org_inner_frame, text=txt, variable=var)
                cb.pack(anchor="w", padx=5, pady=2)
                self.org_vars.append((var, org))

        except Exception as e:
            self.append_log("orgs", f"Error: {e}")
            messagebox.showerror("Error", f"Fetch organizations failed: {e}")

    def select_all_orgs(self):
        for (v, org) in self.org_vars:
            v.set(True)

    def clear_all_orgs(self):
        for (v, org) in self.org_vars:
            v.set(False)

    ####################################################################
    # 3) FETCH DEVICES
    ####################################################################
    def fetch_devices_for_selected_orgs(self):
        for c in self.dev_inner_frame.winfo_children():
            c.destroy()
        self.device_vars.clear()

        sel_orgs = [o for (v,o) in self.org_vars if v.get()]
        if not sel_orgs:
            messagebox.showerror("Error", "No organizations selected.")
            return

        headers = {"User-Agent":USER_AGENT, "Authorization":f"Bearer {self.access_token}"}
        all_devs = []

        for org in sel_orgs:
            org_id   = org["id"]
            org_name = org["name"]
            params = {"organizationId": org_id, "state":"active", "limit":200}
            try:
                r = requests.get(DEVICES_URL, headers=headers, params=params)
                self.append_log("devices", f"GET {r.url} => {r.status_code}\n{r.text}")
                r.raise_for_status()

                j = r.json()
                devs = j.get("items", [])
                for d in devs:
                    d["_orgName"] = org_name
                    all_devs.append(d)
            except Exception as e:
                self.append_log("devices", f"Error fetching from {org_name}: {e}")
                messagebox.showerror("Error", f"Failed to fetch devices from {org_name}: {e}")

        # Show
        for d in all_devs:
            var = tk.BooleanVar(value=False)
            dev_name = d.get("name","Unnamed")
            dev_id   = d.get("id","")
            org_name = d["_orgName"]
            txt = f"({org_name}) - {dev_name} ({dev_id})"
            cb  = ttk.Checkbutton(self.dev_inner_frame, text=txt, variable=var)
            cb.pack(anchor="w", padx=5, pady=2)
            self.device_vars.append((var, d))

    def select_all_devices(self):
        for (v, d) in self.device_vars:
            v.set(True)

    def clear_all_devices(self):
        for (v, d) in self.device_vars:
            v.set(False)

    ####################################################################
    # 4) TRIGGER OPERATION
    ####################################################################
    def trigger_operation_on_selected(self):
        op = self.operation_var.get()
        if op == "checkMissingUpdates":
            return self.handle_check_missing_updates()
        if op == "inventory":
            return self.handle_inventory()

        selected_devs = [d for (v,d) in self.device_vars if v.get()]
        if not selected_devs:
            messagebox.showerror("Error", "No devices selected.")
            return

        msg_str = self.message_var.get().strip()

        total = len(selected_devs)
        success_count = 0
        fail_count    = 0
        chunk_size    = 5

        for i in range(0, total, chunk_size):
            chunk = selected_devs[i:i+chunk_size]
            dev_ids = [x["id"] for x in chunk]
            ok, sc, fc = self._run_op_chunk(op, dev_ids, msg_str)
            success_count += sc
            fail_count    += fc

        if fail_count == 0:
            messagebox.showinfo("Operation Complete",
                f"'{op}' done.\nDevices={total}, success={success_count}, fail=0.\nSee logs.")
        else:
            messagebox.showwarning("Operation Complete with Errors",
                f"'{op}' done.\nDevices={total}, success={success_count}, fail={fail_count}.\nSee logs.")

    def _run_op_chunk(self, op, dev_ids, msg_str):
        """
        Return (ok, success_count, fail_count).
        If "showMessage", we pass the text in 'message'.
        """
        headers = {"User-Agent":USER_AGENT, "Authorization":f"Bearer {self.access_token}", "Content-Type":"application/json"}
        body = {"operation": op, "targets": dev_ids}
        success_count = 0
        fail_count    = 0

        if op == "showMessage" and msg_str:
            body["parameters"] = {"message": msg_str}

        if op == "turnOnFeature":
            val = simpledialog.askinteger("Debug Logging Timeout (min)", "Enter 5–1440:", minvalue=5, maxvalue=1440)
            if val is None:
                for dev_id in dev_ids:
                    self.append_log("operation", f"Device {dev_id} => CANCEL (turnOnFeature input aborted)")
                return (False, 0, len(dev_ids))
            body["parameters"] = {"feature":"debugLogging","turnOnTimeout":val}

        if op == "assignProfile":
            pid = simpledialog.askinteger("Assign Profile", "Enter numeric profileId:")
            if pid is None:
                for dev_id in dev_ids:
                    self.append_log("operation", f"Device {dev_id} => CANCEL (assignProfile input aborted)")
                return (False, 0, len(dev_ids))
            body["parameters"] = {"profileId": pid}

        self.append_log("operation", f"CHUNK op={op}, devices={dev_ids}, body={body}")
        try:
            r = requests.post(OPERATIONS_URL, json=body, headers=headers)
            txt = r.text
            st  = r.status_code
            self.append_log("operation", f"Resp {st}: {txt}")
            r.raise_for_status()

            j = r.json()
            ms = j.get("multistatus", [])
            for item in ms:
                dev_id = item.get("target","")
                s      = item.get("status",0)
                dtl    = item.get("details","")
                if s == 202:
                    success_count += 1
                    self.append_log("operation", f"Device {dev_id} => SUCCESS( st=202 ) details={dtl}")
                else:
                    fail_count    += 1
                    self.append_log("operation", f"Device {dev_id} => FAIL( st={s} ) details={dtl}")
            return (True, success_count, fail_count)
        except Exception as e:
            self.append_log("operation", f"CHUNK error dev_ids={dev_ids}: {e}")
            return (False, 0, len(dev_ids))

    ####################################################################
    # checkMissingUpdates => read missing updates, export CSV with app names
    ####################################################################
    def handle_check_missing_updates(self):
        selected_devs = [d for (v,d) in self.device_vars if v.get()]
        if not selected_devs:
            messagebox.showerror("Error", "No devices selected.")
            return

        upd_win = tk.Toplevel(self)
        upd_win.title("Reading Missing Updates")

        lbl = ttk.Label(upd_win, text="Checking missing software updates for each selected device...")
        lbl.pack(padx=10, pady=10)

        bar = ttk.Progressbar(upd_win, orient="horizontal", length=300, mode="determinate")
        bar.pack(padx=10, pady=5)
        bar["maximum"] = len(selected_devs)
        bar["value"] = 0

        total = len(selected_devs)
        success_count = 0
        fail_count    = 0
        aborted = [False]

        # We'll store (orgName, devName, devID, missingCount, [listOfApps]) for the CSV
        results = []

        def do_abort():
            aborted[0] = True
            upd_win.destroy()

        btn_abort = ttk.Button(upd_win, text="Abort", command=do_abort)
        btn_abort.pack(pady=5)

        upd_win.update()

        for i, dev in enumerate(selected_devs, start=1):
            if aborted[0]:
                break
            dev_id   = dev.get("id","")
            dev_name = dev.get("name","Unnamed")
            org_name = dev.get("_orgName","")
            # read
            missing_count, app_list, ok = self._read_missing_updates_for_device(dev_id)
            if ok:
                success_count += 1
                results.append((org_name, dev_name, dev_id, missing_count, app_list))
            else:
                fail_count += 1

            bar["value"] = i
            lbl.config(text=f"Checking device {i}/{total} ...")
            upd_win.update()
            time.sleep(0.2)

        upd_win.destroy()

        # short summary
        if aborted[0]:
            messagebox.showinfo("Aborted",
                f"Aborted after {i-1} devices.\nSo far success={success_count}, fail={fail_count}.\nCheck logs.")
            return
        else:
            if fail_count == 0:
                messagebox.showinfo("Updates Check Complete",
                    f"Checked missing updates for {total} devices.\nSuccess={success_count}, fail=0.\nCheck logs.")
            else:
                messagebox.showwarning("Updates Check Complete with Errors",
                    f"Checked {total} devices, success={success_count}, fail={fail_count}.\nSee logs for details.")

        # if any device has missingCount > 0, we ask user if they want CSV
        any_missing = any(r[3] > 0 for r in results)
        if any_missing:
            ans = messagebox.askyesno("Export CSV?",
                "We found missing updates on one or more devices.\nExport a CSV grouped by org/device with app names?")
            if ans:
                self._export_missing_updates_csv(results)

    def _read_missing_updates_for_device(self, dev_id):
        """
        Return (missingCount, [appNames], okBoolean).
        We'll parse the 'items' to get the 'name' of each update.
        """
        headers = {
            "User-Agent": USER_AGENT,
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/x-www-form-urlencoded"
        }
        data = {"deviceId": dev_id, "limit":"100"}
        try:
            r = requests.post(MISSING_UPDATES_URL, data=data, headers=headers)
            st = r.status_code
            txt= r.text
            self.append_log("missing_updates", f"Device {dev_id} => st={st}\n{txt}")
            r.raise_for_status()

            j = r.json()
            items = j.get("items", [])
            missing_count = len(items)
            app_list = []
            for it in items:
                # items have fields: "software", "category", "severity" etc
                # We'll gather the name of the software
                sw = it.get("software", {})
                # or it might be just top-level "name" for the patch, let's see doc
                # The doc says:
                #   name, severity, category, software.name, etc
                # We'll prefer "software.name" if present, else "name"
                sw_name = sw.get("name")
                if sw_name:
                    app_list.append(sw_name)
                else:
                    # fallback
                    nm = it.get("name","UnknownApp")
                    app_list.append(nm)

            return (missing_count, app_list, True)
        except Exception as e:
            self.append_log("missing_updates", f"Device {dev_id} => FAIL reading updates: {e}")
            return (0, [], False)

    def _export_missing_updates_csv(self, results):
        """
        results is list of (orgName, devName, devId, missingCount, [appList])
        We'll sort by orgName then devName. Then each device is a row.
        We'll store the app list in a single cell, comma-separated.
        """
        path = filedialog.asksaveasfilename(
            title="Save Missing Updates CSV",
            defaultextension=".csv",
            filetypes=[("CSV File","*.csv"), ("All Files","*.*")]
        )
        if not path:
            return

        sorted_res = sorted(results, key=lambda r: (r[0], r[1]))
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Organization","Device Name","Device ID","Missing Updates Count","Apps Requiring Updates"])
                for (orgN, devN, devID, mc, app_list) in sorted_res:
                    apps_str = ", ".join(app_list) if app_list else ""
                    w.writerow([orgN, devN, devID, mc, apps_str])
            messagebox.showinfo("Export Successful", f"CSV exported to: {path}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not write CSV: {e}")

    ####################################################################
    # inventory => gather hardware info from the device dictionary, export CSV
    ####################################################################
    def handle_inventory(self):
        """
        For each selected device, gather hardware info from the device dictionary
        (like serialNumber, biosVersion, physicalMemoryTotalSize, etc.)
        Then export to CSV grouped by org.
        We'll do a short progress bar, no new API calls (we rely on existing data).
        """
        selected_devs = [d for (v,d) in self.device_vars if v.get()]
        if not selected_devs:
            messagebox.showerror("Error", "No devices selected.")
            return

        inv_win = tk.Toplevel(self)
        inv_win.title("Gathering Inventory")

        lbl = ttk.Label(inv_win, text="Collecting local hardware info from each selected device data...")
        lbl.pack(padx=10, pady=10)

        bar = ttk.Progressbar(inv_win, orient="horizontal", length=300, mode="determinate")
        bar.pack(padx=10, pady=5)
        bar["maximum"] = len(selected_devs)
        bar["value"]   = 0

        total = len(selected_devs)
        aborted = [False]
        results = []

        def do_abort():
            aborted[0] = True
            inv_win.destroy()

        btn_abort = ttk.Button(inv_win, text="Abort", command=do_abort)
        btn_abort.pack(pady=5)

        inv_win.update()

        for i, dev in enumerate(selected_devs, start=1):
            if aborted[0]:
                break
            # gather data from dev
            org_name  = dev.get("_orgName","")
            dev_name  = dev.get("name","")
            dev_id    = dev.get("id","")
            bios      = dev.get("biosVersion","N/A")
            sn        = dev.get("serialNumber","N/A")
            mem_total = dev.get("physicalMemoryTotalSize",0)
            os_info   = dev.get("os",{})
            os_name   = os_info.get("name","")
            os_ver    = os_info.get("version","")
            # possibly more fields if you want

            results.append((org_name, dev_name, dev_id, bios, sn, mem_total, os_name, os_ver))

            bar["value"] = i
            lbl.config(text=f"Processing device {i}/{total} ...")
            inv_win.update()
            time.sleep(0.2)

        inv_win.destroy()

        if aborted[0]:
            messagebox.showinfo("Inventory Aborted", f"Aborted after {i-1} device(s). CSV not generated.")
            return

        # ask user to pick path
        path = filedialog.asksaveasfilename(
            title="Save Inventory CSV",
            defaultextension=".csv",
            filetypes=[("CSV File","*.csv"),("All Files","*.*")]
        )
        if not path:
            return

        # sort by org_name, then dev_name
        results_sorted = sorted(results, key=lambda r: (r[0], r[1]))

        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Organization","Device Name","Device ID","BIOS Version","Serial Number",
                            "Memory (Bytes)","OS Name","OS Version"])
                for row in results_sorted:
                    w.writerow(row)
            messagebox.showinfo("Inventory Export", f"CSV exported to {path}")
        except Exception as e:
            messagebox.showerror("Export Failed", f"Could not write CSV: {e}")

    ####################################################################
    # 5) Device State & Deletion
    ####################################################################
    def update_device_state(self, new_state):
        selected_devs = [d for (v,d) in self.device_vars if v.get()]
        if not selected_devs:
            messagebox.showerror("Error", "No devices selected.")
            return

        if not self._confirm_twice(f"Set these devices to '{new_state}'?"):
            return

        headers = {"User-Agent":USER_AGENT,"Authorization":f"Bearer {self.access_token}","Content-Type":"application/json"}
        chunk_size = 5
        total_count = len(selected_devs)
        success_count=0
        fail_count=0

        for i in range(0, total_count, chunk_size):
            chunk = selected_devs[i:i+chunk_size]
            dev_ids = [x["id"] for x in chunk]
            body = {"state": new_state, "targets": dev_ids}
            self.append_log("device_state", f"PATCH => {body}")
            try:
                r = requests.patch(PATCH_DEVICES_URL, json=body, headers=headers)
                r.raise_for_status()
                j = r.json()
                ms= j.get("multistatus", [])
                for item in ms:
                    st= item.get("status",0)
                    dv= item.get("target","")
                    dt= item.get("details","")
                    if st==200:
                        success_count+=1
                        self.append_log("device_state", f"Device {dv} => SUCCESS {new_state}")
                    else:
                        fail_count+=1
                        self.append_log("device_state", f"Device {dv} => FAIL st={st} details={dt}")
            except Exception as e:
                self.append_log("device_state", f"Error chunk={dev_ids}: {e}")
                fail_count += len(dev_ids)

        if fail_count==0:
            messagebox.showinfo("Done",
                f"'{new_state}' state applied to {total_count} device(s).\n"
                f"Success={success_count}, Fail=0.\nCheck logs.")
        else:
            messagebox.showwarning("Done with Errors",
                f"'{new_state}' partial.\n"
                f"Success={success_count}, Fail={fail_count}.\nCheck logs.")

    def delete_devices(self):
        selected_devs = [d for (v,d) in self.device_vars if v.get()]
        if not selected_devs:
            messagebox.showerror("Error", "No devices selected.")
            return
        if not self._confirm_twice("Are you sure you want to DELETE these device(s)?\nIrreversible!"):
            return

        headers={"User-Agent":USER_AGENT,"Authorization":f"Bearer {self.access_token}"}
        chunk_size=20
        total_count=len(selected_devs)
        success_count=0
        fail_count=0

        for i in range(0, total_count, chunk_size):
            chunk= selected_devs[i:i+chunk_size]
            dev_ids= [x["id"] for x in chunk]
            qs= "&".join([f"deviceId={did}" for did in dev_ids])
            url= f"{DELETE_DEVICES_URL}?{qs}"
            self.append_log("delete_devices", f"DELETE => {url}")
            try:
                r= requests.delete(url, headers=headers)
                r.raise_for_status()
                j= r.json()
                del_list= j.get("devices",[])
                success_count+= len(del_list)
            except Exception as e:
                self.append_log("delete_devices", f"Error chunk={dev_ids}: {e}")
                fail_count+= len(dev_ids)

        if fail_count==0:
            messagebox.showinfo("Deleted Devices",
                f"Deleted {total_count} device(s). success={success_count}, fail=0.\nCheck logs.")
        else:
            messagebox.showwarning("Deleted with Errors",
                f"Deleted {total_count} devices, success={success_count}, fail={fail_count}.\nCheck logs.")

        self.fetch_devices_for_selected_orgs()

    ####################################################################
    # HELPER: double confirm
    ####################################################################
    def _confirm_twice(self, msg):
        ans1 = messagebox.askyesno("Confirm", msg)
        if not ans1:
            return False
        ans2 = messagebox.askyesno("Confirm Again", "Are you absolutely sure?")
        return ans2

# ---------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------
if __name__=="__main__":
    app=WithSecureBatchApp()
    app.mainloop()
