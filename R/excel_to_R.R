#' Convert excel object to data.frame
#'
#' This function is used to excel data to data.frame. Can be used in shiny app to convert input json to data.frame
#' @export
#' @param excelObj the json data retuned from excel table
#' @examples
#'if(interactive()){
#'  library(shiny)
#'  library(excelR)
#'  shinyApp(
#'    ui = fluidPage(excelOutput("table")),
#'    server = function(input, output, session) {
#'      output$table <-
#'        renderExcel(excelTable(data = head(iris)))
#'      observeEvent(input$table,{
#'        print(excel_to_R(input$table))
#'      })
#'    }
#'  )
#'}

excel_to_R <- function(excelObj) {

   if (!is.null(excelObj)) {

      data       <- excelObj$data
      colHeaders <- unlist(excelObj$colHeaders)
      colType    <- unlist(excelObj$colType)

      dataOutput <- data.table::rbindlist(data)

      setnames(dataOutput, colHeaders)

      # JSON is character for all types except logical

      # Change clandar variables to date
      if (any(colType %in% "calendar")){

         dataOutput[,
                    (names(dataOutput)[colType %in% "calendar"]) :=
                    lapply(.SD, FUN = function(i) {

                        as.Date(i, format = "%Y-%m-%d")

                    }),
                    .SDcols = names(dataOutput)[colType %in% "calendar"]
                    ]

      }

      if (any(colType %in% c("dropdown", "text"))) {

         c.check <- c("dropdown", "text")

         dataOutput[,
                    (names(dataOutput)[colType %in% c.check]) := lapply(.SD, FUN = function(i) {

                        ifelse(i == "", NA_character_, i)

                    }),
                    .SDcols = names(dataOutput)[colType %in% c.check]
                    ]

      }

      if (any(colType %in% "numeric")) {

         c.check <- "numeric"

         dataOutput[,
                    (names(dataOutput)[colType %in% c.check]) :=
                    lapply(.SD, FUN = function(i) {

                        as.numeric(i)

                    }),
                    .SDcols = names(dataOutput)[colType %in% c.check]
                    ]

      }

      # return
      dataOutput
   }

}
