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

namespace eval intranet_openoffice:: {}

ad_proc -public intranet_openoffice::spreadsheet {
    {-sql:required}
    {-view_name:required}
    {-ods_file ""}
    {-table_name ""}
    {-variable_set ""}
    {-code ""}
    {-output_filename:required}
} {
    Takes a SQL statement and a view_name to create a Spreadsheet for each column in the view with all the rows from the SQL
    
    @param sql A SQL statement which is used to get each row of the spreadsheet
    @param view_name Name of the dynfield view for which to generate the Spreadsheet
    @param object_type If the object_type is provided we can try to figure out which widget to use for including the column. Only works if there is no extra select involved.
    @param variable_set ns_set of variables we need locally.
    @param code Code which is executed to set or ammend variables.
} {

    # Check if we have the table.ods file in the proper place
    if {$ods_file eq ""} {
        set ods_file "[acs_package_root_dir "intranet-openoffice"]/templates/table.ods"
    }
    if {![file exists $ods_file]} {
        ad_return_error "Missing ODS" "We are missing your ODS file $ods_file . Please make sure it exists"
    }
    
    # Get the "view" (=list of columns to show)
    set view_id [util_memoize [list db_string get_view_id "select view_id from im_views where view_name = '$view_name'" -default 0]]
    if {0 == $view_id} {
        ad_return_error Error "intranet_openoffice::spreadsheet: We didn't find view_name=$view_name"
    }

    if {$variable_set ne ""} {
        ad_ns_set_to_tcl_vars -duplicates ignore $variable_set
    }

    # The table_name is not allowed to contain any quotes
    regsub -all {\"} $table_name {'} table_name

    # ---------------------- Get Columns ----------------------------------
    # Define the column headers and column contents that
    # we want to show:
    #
    
    set variables [list]
    set column_sql "
	select	*
	from	im_view_columns
	where	view_id=:view_id
		and group_id is null
	order by sort_order
    "

    set __column_defs ""
    set __header_defs ""
    
    db_foreach column_list_sql $column_sql {
        
        # We need to check the visibility on the calling procedure....
        if {"" == $visible_for || [eval $visible_for]} {
            if {$variable_name ne ""} {
                switch $datatype {
                    date {
                        append __column_defs "<table:table-column table:style-name=\"co2\" table:default-cell-style-name=\"ce1\"/>\n"
                    }
                    currency {
                        append __column_defs "<table:table-column table:style-name=\"co2\" table:default-cell-style-name=\"ce2\"/>\n"
                    }
                    float {
                        append __column_defs "<table:table-column table:style-name=\"co2\" table:default-cell-style-name=\"ce5\"/>\n"
                    }
                    percentage {
                        append __column_defs "<table:table-column table:style-name=\"co2\" table:default-cell-style-name=\"ce6\"/>\n"
                    }
                    textarea {
                        # Don't shrink to fit textareas. But Autobreak them
                        # this is done using style ce4
                        append __column_defs "<table:table-column table:style-name=\"co3\" table:default-cell-style-name=\"ce4\"/>\n"
                    }
                    default {
                        # style ce3 is set to "shrink to fit", so the size of
                        # the font automatically decreases if needed
                        append __column_defs "<table:table-column table:style-name=\"co1\" table:default-cell-style-name=\"ce3\"/>\n"
                    }
                }

                # Localize the string
                set key [lang::util::suggest_key $column_name]
                set column_name [lang::message::lookup "" intranet-core.$key $column_name]

                append __header_defs " <table:table-cell office:value-type=\"string\"><text:p>$column_name</text:p></table:table-cell>\n"
                set datatype_arr($variable_name) $datatype
                lappend variables $variable_name
            }
        }
    }
    
    set __output $__column_defs
    
    # Set the first row
    append __output "<table:table-row table:style-name=\"ro1\">\n$__header_defs</table:table-row>\n"
    
    # Now create the single rows for each Object
    db_foreach elements $sql {
        append __output "<table:table-row table:style-name=\"ro1\">\n"
        
        foreach variable $variables {
            set value [set $variable]
            eval $code
            switch $datatype_arr($variable) {
                date {
                    append __output " <table:table-cell office:value-type=\"date\" office:date-value=\"[lc_time_fmt $value %F]\"></table:table-cell>\n"
                }
                currency {
                    append __output " <table:table-cell office:value-type=\"currency\" office:currency=\"EUR\" office:value=\"$value\"></table:table-cell>\n"
                }
                percentage {
                    if {$value ne ""} {
                        set value [expr $value / 100]
                    }
                    append __output "<table:table-cell office:value-type=\"percentage\" office:value=\"$value\"></table:table-cell>"
                }
                float {
                    append __output "<table:table-cell office:value-type=\"float\" office:value=\"$value\"></table:table-cell>"
                }
		category_pretty {
		    set category_pretty [im_category_from_id $value]
		    append __output " <table:table-cell office:value-type=\"string\"><text:p>$category_pretty</text:p></table:table-cell>\n"
		}
                default {
                    append __output " <table:table-cell office:value-type=\"string\"><text:p>$value</text:p></table:table-cell>\n"
                }
            }
        }
        append __output "</table:table-row>\n"
    }

    intranet_oo::parse_content -template_file_path $ods_file -output_filename $output_filename
}

ad_proc -public intranet_openoffice::invoice_pdf {
    {-invoice_id:required}
} {
    Generate a PDF for an invoice and saves it as a CR Item
} {
    # First we need to retrieve the invoice
    set user_id [im_sysadmin_user_default]
    set expiry_date [db_string current_date "select to_char(sysdate, 'YYYY-MM-DD') from dual"]
    set auto_login [im_generate_auto_login -expiry_date $expiry_date -user_id $user_id]
    set invoice_url [export_vars -base "[ad_url]/intranet-invoices/view" -url {invoice_id user_id expiry_date auto_login {pdf_p 1} {render_template_id 1}}]
    set mime_type "application/pdf"
    set invoice_nr [db_string name "select invoice_nr from im_invoices where invoice_id = :invoice_id"]

    set tmp_filename [ns_tmpnam]                                                                                  
    apm_transfer_file -url $invoice_url -output_file_name $tmp_filename

    set item_id [content::item::get_id_by_name -name ${invoice_nr}.pdf -parent_id $invoice_id]
    if {$item_id ne ""} {
	set file_revision_id [cr_import_content -item_id $item_id -creation_user $user_id -title "${invoice_nr}.pdf" $invoice_id $tmp_filename [file size $tmp_filename] "application/pdf" "${invoice_nr}.pdf"]
    } else {
	set file_revision_id [cr_import_content -creation_user $user_id -title "${invoice_nr}.pdf" $invoice_id $tmp_filename [file size $tmp_filename] "application/pdf" "${invoice_nr}.pdf"]
    }	
    
    content::item::set_live_revision -revision_id $file_revision_id
    return $file_revision_id
}

ad_proc -public intranet_openoffice::invoices_pdfs {
    {-invoice_ids:required}
    {-order_by "invoice_nr asc"}
} {
    Returns a JOINED PDF of all the invoices provided
} {
    set filenames [list]
    set output_filename "Invoices.pdf"

    db_foreach invoice_ids "select invoice_id,invoice_nr,acs_objects.last_modified from im_invoices,acs_objects,im_costs c where object_id = invoice_id and cost_id = invoice_id and invoice_id in ([template::util::tcl_to_sql_list $invoice_ids]) order by $order_by" {
	set invoice_item_id [content::item::get_id_by_name -name "${invoice_nr}.pdf" -parent_id $invoice_id]

	if {"" == $invoice_item_id} {
	    set invoice_revision_id [intranet_openoffice::invoice_pdf -invoice_id $invoice_id]
	    set invoice_item_id [content::item::get_id_by_name -name "${invoice_nr}.pdf" -parent_id $invoice_id]
	}
	
	if {![file exists "/tmp/${invoice_nr}.pdf"]} {
	    lappend filenames  [fs::publish_object_to_file_system  -object_id $invoice_item_id -path /tmp]
	} else {
	    lappend filenames "/tmp/${invoice_nr}.pdf"
	}
    }
    if {[catch {intranet_oo::join_pdf -filenames $filenames -no_import} pdf_info]} {
	foreach filename $filenames {
	    exec zip -j /tmp/invoices.zip $filename
	}
	set pdf_info [list "application/zip" "/tmp/invoices.zip"]
	set output_filename "Invoices.zip"
    }

    # Delete the original PDFs
    foreach filename $filenames {
        file delete $filename
    }

    set outputheaders [ns_conn outputheaders]
    ns_set cput $outputheaders "Content-Disposition" "attachment; filename=$output_filename"
    ns_returnfile 200 [lindex $pdf_info 0] [lindex $pdf_info 1]
    file delete [lindex $pdf_info 1]
}
