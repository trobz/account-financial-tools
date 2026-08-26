# Copyright (C) 2016-Today: La Louve (<http://www.lalouve.fr/>)
# @author: La Louve
# Copyright (C) 2020-Today: Druidoo (<https://www.druidoo.io>)
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl.html).

import base64
from io import BytesIO

from xlsxwriter.utility import xl_rowcol_to_cell

from odoo import fields, models
from odoo.tools import image_process
from odoo.tools.misc import format_date


class ReportAccountAssetXlsx(models.AbstractModel):
    _inherit = "report.report_xlsx.abstract"
    _name = "report.report_account_asset_xlsx"
    _description = "Report Account Asset XLSX"

    def generate_xlsx_report(self, workbook, data, objects):
        wizard_rec = objects and objects[0] or objects

        resources = dict(**self._context)
        resources.update(
            {
                "row_pos": 5,
                "workbook": workbook,
            }
        )
        resources.update(self._define_formats(workbook))

        sheet = workbook.add_worksheet()
        resources.update({"sheet": sheet})

        self._set_cells_size(resources)
        self.generate_report_title(resources)

        resources = self.generate_report_general(wizard_rec, resources)

        profile_datas_lst = wizard_rec and wizard_rec.get_profile_datas() or []
        resources.update(
            {
                "summary_column_info": {
                    "value": [],
                    "value_residual": [],
                    "salvage_value": [],
                    "amo_ant": [],
                    "amo_de_lan": [],
                    "cum_amo": [],
                    "val_nette": [],
                }
            }
        )
        for profile_datas in profile_datas_lst:
            resources = self.generate_report_profile(profile_datas, resources)

        resources = self.generate_report_summary(resources)

    def _set_cells_size(self, resources):
        sheet = resources["sheet"]
        sheet.set_default_row(20)
        sheet.set_column("A:Z", None, resources["fm_default"])
        sheet.set_column("A:A", 40)
        sheet.set_column("B:L", 20)

    def generate_report_title(self, resources):
        logo_base64 = base64.b64decode(self.env.user.company_id.logo)
        logo_image = image_process(logo_base64, size=(128, 128))
        image_data = BytesIO(logo_image)
        sheet = resources["sheet"]
        sheet.insert_image(0, 0, "company_logo", {"image_data": image_data})
        sheet.merge_range(
            "B2:K5",
            self.env._("Assets And Depreciation"),
            resources["fm_report_title"],
        )

    def generate_report_general(self, wizard_rec, resources):
        row_pos = resources["row_pos"]
        col_pos = 0
        created_infos = self._get_created_by_info(wizard_rec)

        sheet = resources["sheet"]
        sheet.write_rich_string(
            row_pos,
            col_pos,
            resources["fm_bold"],
            created_infos[0],
            resources["fm_default"],
            created_infos[1],
        )

        row_pos += 2
        from_date_header = self.env._("Opening : ")
        from_date_str = self.env["ir.qweb.field.date"].record_to_html(
            wizard_rec, "from_date", {}
        )
        sheet.write_rich_string(
            row_pos,
            col_pos,
            resources["fm_bold"],
            from_date_header,
            resources["fm_default"],
            from_date_str,
        )
        row_pos += 1
        to_date_header = self.env._("Closing : ")
        to_date_str = self.env["ir.qweb.field.date"].record_to_html(
            wizard_rec, "to_date", {}
        )
        sheet.write_rich_string(
            row_pos,
            col_pos,
            resources["fm_bold"],
            to_date_header,
            resources["fm_default"],
            to_date_str,
        )
        resources.update({"row_pos": row_pos + 1})
        return resources

    def generate_report_profile(self, profile_datas, resources):
        row_pos = resources["row_pos"]
        col_pos = 0

        sheet = resources["sheet"]
        summary_column_info = resources["summary_column_info"]
        profile_name = profile_datas.get("profile_name", "")
        sheet.write(row_pos, col_pos, profile_name, resources["fm_table"])

        for i in range(1, 12):
            col_pos = i
            sheet.write(row_pos, col_pos, "", resources["fm_table"])

        row_pos += 1
        table_infos = self._get_table_infos(resources)
        for col_index, column in enumerate(table_infos.keys(), 0):
            col_pos = col_index
            label_dict = table_infos.get(column)
            label_str = label_dict.get("str")
            cell_format = resources["fm_table_header"]
            sheet.write(row_pos, col_pos, label_str, cell_format)

        row_pos += 1
        profile_data_lines = profile_datas.get("lines", [])
        profile_data_lines_length = len(profile_data_lines)
        start_row_pos = row_pos

        for index in range(profile_data_lines_length + 1):
            is_sub_summary_section = False
            if index < profile_data_lines_length:
                line_data = profile_data_lines[index]
            else:
                is_sub_summary_section = True
                stop_row_pos = row_pos - 1
                line_data = {
                    "name": self.env._("Subtotal"),
                    "value": "=SUM({}:{})",
                    "value_residual": "=SUM({}:{})",
                    "salvage_value": "=SUM({}:{})",
                    "amo_ant": "=SUM({}:{})",
                    "amo_de_lan": "=SUM({}:{})",
                    "cum_amo": "=SUM({}:{})",
                    "val_nette": "=SUM({}:{})",
                }
            for col_index, column in enumerate(table_infos.keys(), 0):
                col_pos = col_index
                cell_value = line_data.get(column, "")

                if isinstance(cell_value, bool):
                    cell_value = cell_value and self.env._("Yes") or self.env._("No")

                cell_type = table_infos.get(column).get("type", "")
                cell_format = table_infos.get(column).get(
                    "format", resources["fm_default"]
                )
                need_to_write_summary_formula = (
                    is_sub_summary_section and cell_type == "formula"
                )

                if is_sub_summary_section:
                    cell_format = resources["fm_table_bold"]

                # Write formula
                if need_to_write_summary_formula:
                    cell_format = resources["fm_table_number_bold"]
                    start_cell = xl_rowcol_to_cell(start_row_pos, col_pos)
                    stop_cell = xl_rowcol_to_cell(stop_row_pos, col_pos)
                    cell_formula_value = cell_value.format(start_cell, stop_cell)
                    sheet.write_formula(
                        row_pos, col_pos, cell_formula_value, cell_format
                    )

                    # append cell to summary section info
                    summary_column_info[column].append(
                        xl_rowcol_to_cell(row_pos, col_pos)
                    )
                else:
                    if column == "date" and not is_sub_summary_section:
                        cell_format = resources["fm_table_date"]
                        cell_value = format_date(self.env, cell_value)

                    sheet.write(row_pos, col_pos, cell_value, cell_format)
            row_pos += 1
        resources.update(
            {
                "row_pos": row_pos + 1,
                "summary_column_info": summary_column_info,
            }
        )
        return resources

    def generate_report_summary(self, resources):
        sheet = resources["sheet"]
        summary_column_info = resources["summary_column_info"]
        row_pos = resources["row_pos"]
        summary_data = {
            "name": self.env._("Total"),
            "value": "=SUM({})",
            "value_residual": "=SUM({})",
            "salvage_value": "=SUM({})",
            "amo_ant": "=SUM({})",
            "amo_de_lan": "=SUM({})",
            "cum_amo": "=SUM({})",
            "val_nette": "=SUM({})",
        }
        cell_format = resources["fm_table_header_dark_grey"]
        table_infos = self._get_table_infos(resources)
        for col_index, column in enumerate(table_infos, 0):
            col_pos = col_index
            cell_value = summary_data.get(column, "")
            if "=SUM" in cell_value:
                cell_format = resources["fm_table_number_bold_dark_grey"]
                cell_formula_value = cell_value.format(
                    ", ".join(summary_column_info[column])
                )
                sheet.write_formula(row_pos, col_pos, cell_formula_value, cell_format)
            else:
                sheet.write(row_pos, col_pos, cell_value, cell_format)
        resources.update(
            {
                "row_pos": row_pos + 1,
                "summary_column_info": summary_column_info,
            }
        )
        return resources

    def _get_created_by_info(self, wizard_rec):
        today = fields.Date.today()
        formated_today_str = format_date(self.env, today)
        created_by_header = self.env._("Created on : ")
        created_uid = wizard_rec._context.get("uid", self.env.uid)
        created_user = self.env["res.users"].browse(created_uid)
        created_by_info = self.env._(
            "%s by %s",
            formated_today_str,
            created_user.name,
        )
        return created_by_header, created_by_info

    def _get_summary_column_info(self):
        return {
            "value": [],
            "value_residual": [],
            "salvage_value": [],
            "amo_ant": [],
            "amo_de_lan": [],
            "cum_amo": [],
            "val_nette": [],
        }

    def _get_table_infos(self, resources):
        return {
            "name": {"str": self.env._("Asset Name")},
            "state": {"str": self.env._("Status")},
            "date": {"str": "Date", "format": resources["fm_table_date"]},
            "value": {
                "str": self.env._("Gross Value"),
                "type": "formula",
                "format": resources["fm_table_number"],
            },
            "salvage_value": {
                "str": self.env._("Salvage Value"),
                "type": "formula",
                "format": resources["fm_table_number"],
            },
            "method": {"str": self.env._("Computation Method")},
            "method_number": {"str": self.env._("Nb. of Years")},
            "prorata": {"str": self.env._("Prorata Temporis")},
            "amo_ant": {
                "str": self.env._("Beg. Acc. Dep."),
                "type": "formula",
                "format": resources["fm_table_number"],
            },
            "amo_de_lan": {
                "str": self.env._("Depreciation Expense"),
                "type": "formula",
                "format": resources["fm_table_number"],
            },
            "cum_amo": {
                "str": self.env._("End. Acc. Dep."),
                "type": "formula",
                "format": resources["fm_table_number"],
            },
            "value_residual": {
                "str": self.env._("Residual Value"),
                "type": "formula",
                "format": resources["fm_table_number"],
            },
        }

    def _create_format(self, workbook, format_dict):
        return workbook.add_format(format_dict)

    def _define_formats(self, workbook):
        fm_values = {}
        # ---------------------------------------------------------------------
        # Common
        # ---------------------------------------------------------------------
        fm_default = {
            "font_size": 10,
            "valign": "vcenter",
            "text_wrap": True,
        }
        c_fm_default = self._create_format(workbook, fm_default)
        fm_values.update({"fm_default": c_fm_default})
        # ---------------------------------------------------------------------
        fm_bold = fm_default.copy()
        fm_bold.update({"bold": True})
        c_fm_bold = self._create_format(workbook, fm_bold)
        fm_values.update({"fm_bold": c_fm_bold})
        # ---------------------------------------------------------------------
        fm_center = fm_default.copy()
        fm_center.update({"align": "center"})
        c_fm_center = self._create_format(workbook, fm_center)
        fm_values.update({"fm_center": c_fm_center})

        # ---------------------------------------------------------------------
        # Report Title
        # ---------------------------------------------------------------------
        fm_report_title = fm_default.copy()
        fm_report_title.update({"bold": True, "align": "center", "font_size": 36})
        c_fm_report_title = self._create_format(workbook, fm_report_title)
        fm_values.update({"fm_report_title": c_fm_report_title})
        # ---------------------------------------------------------------------
        fm_title_table = fm_default.copy()
        fm_title_table.update({"bold": True, "align": "center"})
        c_fm_title_table = self._create_format(workbook, fm_title_table)
        fm_values.update({"fm_title_table": c_fm_title_table})

        # ---------------------------------------------------------------------
        # Table format
        # ---------------------------------------------------------------------
        fm_table = fm_default.copy()
        fm_table.update(
            {
                "font_size": 11,
                "bold": True,
                "align": "vcenter",
                "bg_color": "#0070c0",
                "font_color": "#ffffff",
            }
        )
        c_fm_table = self._create_format(workbook, fm_table)
        fm_values.update({"fm_table": c_fm_table})
        # ---------------------------------------------------------------------
        fm_table_header = fm_table.copy()
        fm_table_header.update(
            {
                "font_size": 10,
                "bg_color": "#d9d9d9",
                "font_color": "#000000",
            }
        )
        c_fm_table_header = self._create_format(workbook, fm_table_header)
        fm_values.update({"fm_table_header": c_fm_table_header})
        # ---------------------------------------------------------------------
        fm_table_bold = fm_table.copy()
        fm_table_bold.update(
            {
                "font_size": 10,
                "bg_color": "#d9d9d9",
                "font_color": "#000000",
            }
        )
        c_fm_table_bold = self._create_format(workbook, fm_table_bold)
        fm_values.update({"fm_table_bold": c_fm_table_bold})
        # ---------------------------------------------------------------------
        fm_table_number = fm_table.copy()
        fm_table_number.update(
            {
                "font_size": 10,
                "bold": False,
                "num_format": "#,##0.00",
            }
        )
        c_fm_table_number = self._create_format(workbook, fm_table_number)
        fm_values.update({"fm_table_number": c_fm_table_number})
        # ---------------------------------------------------------------------
        fm_table_number_bold = fm_table.copy()
        fm_table_number_bold.update(
            {
                "font_size": 10,
                "num_format": "#,##0.00",
                "bg_color": "#d9d9d9",
                "font_color": "#000000",
            }
        )
        c_fm_table_number_bold = self._create_format(workbook, fm_table_number_bold)
        fm_values.update({"fm_table_number_bold": c_fm_table_number_bold})
        # ---------------------------------------------------------------------
        fm_table_date = fm_default.copy()
        fm_table_date.update(
            {
                "font_size": 10,
            }
        )
        c_fm_table_date = self._create_format(workbook, fm_table_date)
        fm_values.update({"fm_table_date": c_fm_table_date})
        # ---------------------------------------------------------------------
        fm_table_header_dark_grey = fm_table_header.copy()
        fm_table_header_dark_grey.update(
            {
                "bg_color": "#808080",
                "font_color": "#000000",
            }
        )
        c_fm_table_header_dark_grey = self._create_format(
            workbook, fm_table_header_dark_grey
        )
        fm_values.update(
            {
                "fm_table_header_dark_grey": c_fm_table_header_dark_grey,
            }
        )
        # ---------------------------------------------------------------------
        fm_table_number_bold_dark_grey = fm_table_number_bold.copy()
        fm_table_number_bold_dark_grey.update(
            {
                "bg_color": "#808080",
                "font_color": "#000000",
            }
        )
        c_fm_table_number_bold_dark_grey = self._create_format(
            workbook, fm_table_number_bold_dark_grey
        )
        fm_values.update(
            {
                "fm_table_number_bold_dark_grey": c_fm_table_number_bold_dark_grey,
            }
        )
        return fm_values
