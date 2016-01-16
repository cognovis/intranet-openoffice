# Copyright (c) 2011, cognovís GmbH, Hamburg, Germany
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see
# <http://www.gnu.org/licenses/>.
#

ad_library {
    
    Procedures for composing correctly prepared content for OpenOffice integration
    
    @author Malte Sussdorff (malte.sussdorff@cognovis.de)
    @creation-date 2011-03-28
}

# ---------------------------------------------------------------
# Projects
# ---------------------------------------------------------------

ad_proc -public -callback im_projects_index_filter -impl intranet-openoffice-spreadsheet {
    {-form_id:required}
} {
    Add the filter for the view_type
} {
    if {[im_permission [ad_conn user_id] "oo_download_projects"]} {
        uplevel {
            set view_type_options [concat $view_type_options [list [list Excel xls]] [list [list Openoffice ods]] [list [list PDF pdf]]]
        }
    }
}

ad_proc -public -callback im_projects_index_before_render -impl intranet-openoffice-spreadsheet {
    {-view_name:required}
    {-view_type:required}
    {-sql:required}
    {-table_header ""}
    {-variable_set ""}
} {
    Depending on the view_type return a spreadsheet in Excel / Openoffice or PDF
} {
    if {[im_permission [ad_conn user_id] "oo_download_projects"]} {

        # Only execute for view types which are supported
        if {[lsearch [list xls pdf ods] $view_type] > -1} {
            intranet_openoffice::spreadsheet -view_name $view_name -sql $sql -output_filename "projects.$view_type" -table_name "$table_header" -variable_set $variable_set
            ad_script_abort
        }
    }
}


# ---------------------------------------------------------------
# Companies
# ---------------------------------------------------------------


ad_proc -public -callback im_companies_index_filter -impl intranet-openoffice-spreadsheet {
    {-form_id:required}
} {
    Add the filter for the view_type
} {
    if {[im_permission [ad_conn user_id] "oo_download_companies"]} {
    
        uplevel {
            set view_type_options [concat $view_type_options [list [list Excel xls]] [list [list Openoffice ods]] [list [list PDF pdf]]]
        }
    }
}


ad_proc -public -callback im_companies_index_before_render -impl intranet-openoffice-spreadsheet {
    {-view_name:required}
    {-view_type:required}
    {-sql:required}
    {-table_header ""}
    {-variable_set ""}
} {
    Depending on the view_type return a spreadsheet in Excel / Openoffice or PDF
} {
    if {[im_permission [ad_conn user_id] "oo_download_companies"]} {
    
        # Only execute for view types which are supported
        if {[lsearch [list xls pdf ods] $view_type] > -1} {
            intranet_openoffice::spreadsheet -view_name $view_name -sql $sql -output_filename "projects.$view_type" -table_name "$table_header" -variable_set $variable_set
            ad_script_abort
        }
    }
}



# ---------------------------------------------------------------
# Tasks
# ---------------------------------------------------------------

ad_proc -public -callback im_timesheet_tasks_index_filter -impl intranet-openoffice-spreadsheet {
    {-form_id:required}
} {
    Add the filter for the view_type
} {
    if {[im_permission [ad_conn user_id] "oo_download_tasks"]} {

        uplevel {
            set view_type_options [concat $view_type_options [list [list Excel xls]] [list [list Openoffice ods]] [list [list PDF pdf]]]
        }
    }
}

ad_proc -public -callback im_timesheet_task_list_before_render -impl intranet-openoffice-spreadsheet {
    {-view_name:required}
    {-view_type:required}
    {-sql:required}
    {-table_header ""}
} {
    Depending on the view_type return a spreadsheet in Excel / Openoffice or PDF
} {
    if {[im_permission [ad_conn user_id] "oo_download_tasks"]} {

        # Only execute for view types which are supported
        if {[lsearch [list xls pdf ods] $view_type] > -1} {
            intranet_openoffice::spreadsheet -view_name $view_name -sql $sql -output_filename "tasks.$view_type" -table_name "$table_header"
            ad_script_abort
        }
    }
}


# ---------------------------------------------------------------
# Invoices
# ---------------------------------------------------------------

ad_proc -public -callback im_invoices_after_create -impl intranet-openoffice-pdf-invoice {
    {-object_type:required}
    {-object_id:required}
    {-status_id:required}
    {-type_id:required}
} {
    Generate a PDF for the created invoice document and attach the PDF to the invoice

    Use the content repository for this and make sure you create new revisions not new files.
} {
}

ad_proc -public -callback im_invoices_index_before_render -impl intranet-openoffice-spreadsheet {
    {-view_name:required}
    {-view_type:required}
    {-sql:required}
    {-table_header ""}
    {-variable_set ""}
} {
    Depending on the view_type return a spreadsheet in Excel / Openoffice or PDF
} {
     if {[im_permission [ad_conn user_id] "oo_download_invoices"]} {

        upvar 1 cost_type_id invoice_type_id
        set invoice_type [im_category_from_id $invoice_type_id]
        # Only execute for view types which are supported
        if {[lsearch [list xls pdf ods] $view_type] > -1} {
            intranet_openoffice::spreadsheet -view_name $view_name -sql $sql -output_filename "${invoice_type}-list.$view_type" -table_name "$table_header" -variable_set $variable_set
            ad_script_abort
        }
    }
}


# ---------------------------------------------------------------
# Timesheets
# ---------------------------------------------------------------


ad_proc -public -callback im_timesheet_report_filter -impl intranet-openoffice-spreadsheet {
    {-form_id:required}
} {
    Add the filter for the output_format
} {
    if {[im_permission [ad_conn user_id] "oo_download_timesheets"]} {

        uplevel {
            set output_format_options [concat $output_format_options [list [list Excel xls]] [list [list Openoffice ods]] [list [list PDF pdf]]]
        }
    }
}

ad_proc -public -callback im_timesheet_report_before_render -impl intranet-openoffice-spreadsheet {
    {-view_name:required}
    {-view_type:required}
    {-sql:required}
    {-table_header ""}
    {-variable_set ""}
} {
    Depending on the view_type return a spreadsheet in Excel / Openoffice or PDF
} {
    if {[im_permission [ad_conn user_id] "oo_download_timesheets"]} {

        # Only execute for view types which are supported
        if {[lsearch [list xls pdf ods] $view_type] > -1} {
            intranet_openoffice::spreadsheet -view_name $view_name -sql $sql -output_filename "timesheet.$view_type" -table_name "$table_header" -variable_set $variable_set
            ad_script_abort
        }
    }
}

# ---------------------------------------------------------------
# Users
# ---------------------------------------------------------------

ad_proc -public -callback im_users_index_filter -impl intranet-openoffice-spreadsheet {
    {-form_id:required}
} {
    Add the filter for the view_type
} {
    if {[im_permission [ad_conn user_id] "oo_download_users"]} {
        uplevel {
            set view_type_options [concat $view_type_options [list [list Excel xls]] [list [list Openoffice ods]] [list [list PDF pdf]]]
        }
    }
}

ad_proc -public -callback im_users_index_before_render -impl intranet-openoffice-spreadsheet {
    {-view_name:required}
    {-view_type:required}
    {-sql:required}
    {-table_header ""}
    {-variable_set ""}
} {
    Depending on the view_type return a spreadsheet in Excel / Openoffice or PDF
} {
    if {[im_permission [ad_conn user_id] "oo_download_users"]} {

        # Only execute for view types which are supported
        if {[lsearch [list xls pdf ods] $view_type] > -1} {
            intranet_openoffice::spreadsheet -view_name $view_name -sql $sql -output_filename "users.$view_type" -table_name "$table_header" -variable_set $variable_set
            ad_script_abort
        }
    }
}