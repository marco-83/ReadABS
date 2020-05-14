# library(devtools)
# install_github("kevinushey/reticulate")
# import_xls <- NULL
# main_xls <- NULL
# xlrd <- NULL

# Load the module and create dummy objects from it, all of which are NULL
ABStable <- reticulate::import_from_path(
  "ABStable",
  file.path("inst", "python")
)
for (obj in names(ABStable)) {
  assign(obj, NULL)
}
rm(ABStable)

create_df_xls <- reticulate::import_from_path(
  "create_df_xls",
  file.path("inst", "python")
)
for (obj in names(create_df_xls)) {
  assign(obj, NULL)
}
rm(create_df_xls)

ABStable_xlsx <- reticulate::import_from_path(
  "ABStable_xlsx",
  file.path("inst", "python")
)
for (obj in names(ABStable_xlsx)) {
  assign(obj, NULL)
}
rm(ABStable_xlsx)

create_df_xlsx <- reticulate::import_from_path(
  "create_df_xlsx",
  file.path("inst", "python")
)
for (obj in names(create_df_xlsx)) {
  assign(obj, NULL)
}
rm(create_df_xlsx)


.onLoad <- function(libname, pkgname) {
  reticulate::configure_environment(pkgname)
  # reticulate::py_install("openpyxl")
  # reticulate::py_install("xlsxwriter")

  pkg_ns_env <- parent.env(environment())
  packages <- c("pandas", "xlrd", "openpyxl", "numpy", "XlsxWriter")
  for (package in packages) {
    if (reticulate::py_module_available(package) == FALSE) {
      reticulate::py_install(package)
    }
    #assign(package,reticulate::import(package),pkg_ns_env)
  }

  #reticulate::configure_environment(pkgname)
  #xlrd <<- reticulate::import("xlrd", delay_load = FALSE)
  #xlrd <- reticulate::import("xlrd")
  ABStable <- reticulate::import_from_path(
    module = "ABStable",
    path = system.file("python", package = packageName())
  )
  # assignInMyNamespace(...) is meant for namespace manipulation
  for (obj in names(ABStable)) {
    assignInMyNamespace(obj, ABStable[[obj]])
    assign(obj,obj,pkg_ns_env)
  }
  # ABStable
  # module1 <- reticulate::import_from_path(module = "import_xls",
  #                                            path =system.file("python",
  #                                                              package = packageName()))
  create_df_xls <- reticulate::import_from_path(module = "create_df_xls",
                                          path = system.file("python",
                                                             package = packageName()))

  for (obj in names(create_df_xls)) {
    assignInMyNamespace(obj, create_df_xls[[obj]])
    assign(obj,obj,pkg_ns_env)
  }

  ABStable_xlsx <- reticulate::import_from_path(
    module = "ABStable_xlsx",
    path = system.file("python", package = packageName())
  )
  # assignInMyNamespace(...) is meant for namespace manipulation
  for (obj in names(ABStable_xlsx)) {
    assignInMyNamespace(obj, ABStable_xlsx[[obj]])
    assign(obj,obj,pkg_ns_env)
  }

  create_df_xlsx <- reticulate::import_from_path(
    module = "create_df_xlsx",
    path = system.file("python", package = packageName())
  )
  # assignInMyNamespace(...) is meant for namespace manipulation
  for (obj in names(create_df_xlsx)) {
    assignInMyNamespace(obj, create_df_xlsx[[obj]])
    assign(obj,obj,pkg_ns_env)
  }


  # lapply(names(module3), function(name) assign(name, module3[[name]], pkg_ns_env))
  #import_xls <<- module1$import_xls
  main_xls <<- create_df_xls$main_xls
  main_xlsx <<- create_df_xlsx$main_xlsx
  #define_table <<- ABStable$define_table
  # import_spreadsheet <<- module3$import_spreadsheet

}
# packages <- c("itertools", "xlrd", "copy", "openpyxl", "operator", "xlsxwriter")
# for (package in packages) {
#   if (reticulate::py_module_available(package) == FALSE) {
#     reticulate::py_install(package)
#   }
#   assign(package,reticulate::import(package),.GlobalEnv)
# }
# reticulate::py_module_available("itertools")
# xlrd <- reticulate::import("xlrd")
# copy <- reticulate::import("copy")
# itertools <- reticulate::import("itertools")
# #reticulate::py_install("openpyxl")
# openpyxl <- reticulate::import("openpyxl")
# operator <- reticulate::import("operator")
#xlsxwriter <- reticulate::import("xlsxwriter")

# module1 <- reticulate::import_from_path(module = "import_xls",
#                                         path = system.file("inst/python",
#                                                            package = "<mypackage>"))
# module2 <- reticulate::import_from_path(module = "ABStable.py",
#                                         path = system.file("python",
#                                                            package = "<mypackage>"))
# import_xls <<- module1$import_xls
# main_xls <<- module1$main_xls

# reticulate::source_python("import_xls.py")
# reticulate::source_python("ABStable.py")
# reticulate::source_python("create_df_xls.py")

# path <- system.file("python", package = "<pkg>")
# module <- reticulate::import_from_path(<module>, path = path)

#' Tidy an xls or xlsx file
#'
#' This function puts xls or xlsx files in a standardised 'long' format, retaining descriptive information in the spreadsheet. It is calibrated to work on Australian Bureau of Statistics, 'ABS' spreadsheets.
#'
#' @param xl_workbook Path to the input file
#' @param allowed_blank_rows Number of blank rows between tables, options are 1 or 2. Defaults to 1.
#' @param spreadsheet_type Spreadsheet type. Options are 'Time series', 'Data cube', 'Census'. Defaults to 'Data cube'.
#' @return A list containing the cleaned data and the tab, row and column information on where the data was found
#' @export
tidy_ABS <- function(xl_workbook, allowed_blank_rows=1,
                    spreadsheet_type="Data cube") {

  if (!file.exists(xl_workbook)) {
    stop("Invalid file path")
  }
  if (!(allowed_blank_rows %in% c(1, 2))) {
    stop("allowed_blank_rows options are 1 or 2") }

  if (!(spreadsheet_type %in% c("Data cube", "Time series", "Census"))) {
    stop("spreadsheet_type options are 'Data cube', 'Time series' or 'Census") }

  # check file extenstion
  ex <- strsplit(xl_workbook, split="\\.")[[1]][-1]
  if (!(ex %in% c("xls", "xlsx"))) stop("Only .xls and .xlsx file types are allowed")

  if (ex == "xls") {
    output <- main_xls(excel_workbook=xl_workbook, allowed_blank_rows=allowed_blank_rows,
                       spreadsheet_type=spreadsheet_type)
  }
  else if (ex == "xlsx") {
    output <- main_xlsx(excel_workbook=xl_workbook, allowed_blank_rows=allowed_blank_rows,
                        spreadsheet_type=spreadsheet_type)
  }
  output

}

#tidy_ABS(xl_workbook = "81550do005_201718.xls")
# tidy_ABS(xl_workbook = "5204001_key_national_aggregates_xlsx.xlsx", spreadsheet_type = "Time series")

# Imports:
#   reticulate
#vignette("python_dependencies")
#devtools::document()
#devtools::install()
#library(ReadABS)
# git add .
# git commit -m "Initial commit"
#git remote add origin https://github.com/marco-83/ReadABS
# git push -u origin master
