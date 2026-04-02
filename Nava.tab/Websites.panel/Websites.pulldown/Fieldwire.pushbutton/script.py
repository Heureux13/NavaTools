# -*- coding: utf-8 -*-
# ======================================================================
"""Copyright (c) 2025 Jose Francisco Nava Perez. All rights reserved.

This code and associated documentation files may not be copied, modified,
distributed, or used in any form without the prior written permission of
the copyright holder."""
# ======================================================================

from pyrevit import script
from System.Windows.Forms import Form, Label, Button, DialogResult, ComboBox, TextBox
import webbrowser
import json
import os
import clr
clr.AddReference("System.Windows.Forms")


# Button info
# ===================================================
__title__ = "Fieldwire"
__doc__ = """
Shortcut to the Fieldwire website
"""

# Variables
# ==================================================
output = script.get_output()

# Configuration file to store project names and IDs
# Stored in user's AppData so each user has their own copy (not shared via git)
CONFIG_FILE = os.path.join(os.path.expanduser("~"), "AppData", "Roaming", "NavaTools", "fieldwire_projects.txt")

# Create the directory if it doesn't exist
CONFIG_DIR = os.path.dirname(CONFIG_FILE)
if not os.path.exists(CONFIG_DIR):
    try:
        os.makedirs(CONFIG_DIR)
    except:
        pass

# Class
# =====================================================================

class FieldwireForm(Form):
    def __init__(self):
        Form.__init__(self)
        self.Text = "Open Fieldwire"
        self.Width = 500
        self.Height = 350
        self.projects = self.load_projects()

        # Title label
        lbl_title = Label()
        lbl_title.Text = "Fieldwire - Select a project:"
        lbl_title.Top = 20
        lbl_title.Left = 20
        lbl_title.Width = 450
        lbl_title.Height = 20
        self.Controls.Add(lbl_title)

        # Project dropdown
        self.combo_projects = ComboBox()
        self.combo_projects.Top = 45
        self.combo_projects.Left = 20
        self.combo_projects.Width = 450
        for project_name in sorted(self.projects.keys()):
            self.combo_projects.Items.Add(project_name)
        self.Controls.Add(self.combo_projects)

        # Add project button
        btn_add = Button()
        btn_add.Text = "Add New Project"
        btn_add.Top = 75
        btn_add.Left = 20
        btn_add.Width = 150
        btn_add.Click += lambda s, e: self.add_project()
        self.Controls.Add(btn_add)

        # Remove project button
        btn_remove = Button()
        btn_remove.Text = "Remove Project"
        btn_remove.Top = 75
        btn_remove.Left = 200
        btn_remove.Width = 150
        btn_remove.Click += lambda s, e: self.remove_project()
        self.Controls.Add(btn_remove)

        # Separator
        lbl_sep = Label()
        lbl_sep.Text = "Select page to open:"
        lbl_sep.Top = 115
        lbl_sep.Left = 20
        lbl_sep.Width = 450
        lbl_sep.Height = 20
        self.Controls.Add(lbl_sep)

        # Projects button
        btn_projects = Button()
        btn_projects.Text = "Plans"
        btn_projects.Top = 145
        btn_projects.Left = 20
        btn_projects.Width = 100
        btn_projects.Click += lambda s, e: self.open_plans()
        self.Controls.Add(btn_projects)

        # Submittals button
        btn_submittals = Button()
        btn_submittals.Text = "Submittals"
        btn_submittals.Top = 145
        btn_submittals.Left = 130
        btn_submittals.Width = 100
        btn_submittals.Click += lambda s, e: self.open_submittals()
        self.Controls.Add(btn_submittals)

        # Specs button
        btn_specs = Button()
        btn_specs.Text = "Specs"
        btn_specs.Top = 145
        btn_specs.Left = 240
        btn_specs.Width = 100
        btn_specs.Click += lambda s, e: self.open_specs()
        self.Controls.Add(btn_specs)

        # Tasks button
        btn_tasks = Button()
        btn_tasks.Text = "Tasks"
        btn_tasks.Top = 145
        btn_tasks.Left = 350
        btn_tasks.Width = 100
        btn_tasks.Click += lambda s, e: self.open_tasks()
        self.Controls.Add(btn_tasks)

        # Close button
        btn_close = Button()
        btn_close.Text = "Close"
        btn_close.Top = 185
        btn_close.Left = 20
        btn_close.Width = 450
        btn_close.DialogResult = DialogResult.Cancel
        self.Controls.Add(btn_close)
        self.CancelButton = btn_close

    def load_projects(self):
        """Load saved projects from config file"""
        projects = {}
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    projects = json.load(f)
            except:
                pass
        return projects

    def save_projects(self):
        """Save projects to config file"""
        with open(CONFIG_FILE, 'w') as f:
            json.dump(self.projects, f, indent=2)

    def add_project(self):
        """Add a new project"""
        # Get project name
        project_name = self.show_input_dialog("Enter project name:", "Add Project")
        if not project_name:
            return
        
        # Get project ID
        project_id = self.show_input_dialog("Enter Fieldwire project ID:", "Add Project")
        if not project_id:
            return
        
        self.projects[project_name] = project_id
        self.save_projects()
        self.refresh_combo()

    def remove_project(self):
        """Remove selected project"""
        if self.combo_projects.SelectedIndex >= 0:
            project_name = self.combo_projects.SelectedItem
            del self.projects[project_name]
            self.save_projects()
            self.refresh_combo()

    def refresh_combo(self):
        """Refresh dropdown list"""
        self.combo_projects.Items.Clear()
        for project_name in sorted(self.projects.keys()):
            self.combo_projects.Items.Add(project_name)

    def show_input_dialog(self, prompt, title):
        """Show input dialog"""
        dlg = Form()
        dlg.Text = title
        dlg.Width = 400
        dlg.Height = 150
        
        lbl = Label()
        lbl.Text = prompt
        lbl.Top = 20
        lbl.Left = 20
        lbl.Width = 350
        dlg.Controls.Add(lbl)
        
        txt = TextBox()
        txt.Top = 50
        txt.Left = 20
        txt.Width = 350
        dlg.Controls.Add(txt)
        
        btn_ok = Button()
        btn_ok.Text = "OK"
        btn_ok.Top = 85
        btn_ok.Left = 20
        btn_ok.Width = 80
        btn_ok.DialogResult = DialogResult.OK
        dlg.Controls.Add(btn_ok)
        dlg.AcceptButton = btn_ok
        
        if dlg.ShowDialog() == DialogResult.OK:
            return txt.Text
        return None

    def get_selected_project_id(self):
        """Get ID of selected project"""
        if self.combo_projects.SelectedIndex >= 0:
            project_name = self.combo_projects.SelectedItem
            return self.projects.get(project_name)
        return None

    def _get_url(self, base_path):
        """Build URL with project ID"""
        project_id = self.get_selected_project_id()
        if project_id:
            return "https://app.fieldwire.com/projects/{}/{}".format(project_id, base_path)
        else:
            return "https://app.fieldwire.com/index/{}".format(base_path)

    def open_plans(self):
        url = self._get_url("plans")
        webbrowser.open(url)

    def open_submittals(self):
        url = self._get_url("files")
        webbrowser.open(url)

    def open_specs(self):
        url = self._get_url("specifications")
        webbrowser.open(url)

    def open_tasks(self):
        url = self._get_url("tasks")
        webbrowser.open(url)


# Main Code
# ==========================================================================================
try:
    form = FieldwireForm()
    form.ShowDialog()
    # output.print_md("✅ Fieldwire opened in your browser")

except Exception as e:
    output.print_md("❌ Error: {}".format(str(e)))
    import traceback
    output.print_md("```\n{}\n```".format(traceback.format_exc()))