  HTMLWidgets.widget({
    name: "jexcel",

    type: "output",

    factory: function(el, width, height) {
      var elementId = el.id;
      var container = document.getElementById(elementId);
      var excel = null;

      return {
        renderValue: function(params) {
          var rowHeight = params.hasOwnProperty("rowHeight") ? params.rowHeight : undefined;
          var showToolbar = params.hasOwnProperty("showToolbar")? params.showToolbar: false;
          var dateFormat = params.hasOwnProperty("dateFormat")? params.dateFormat: "DD/MM/YYYY";
          var otherParams = {};

          Object.keys(params).forEach(function(ky) {

            if(
               ky !== "dateFormat" &&
               ky !== "rowHeight" &&
               ky !== "otherParams"
               ) {

              // Check if the key is columns and check if the type is calendar, if yes add the date format
              if(ky === "columns" && dateFormat !== "DD/MM/YYYY"){

                otherParams[ky] = params[ky].map(function(column){

                  if(column.type === "calendar"){

                    column.options = {format: dateFormat}

                  }

                  return column

                })

                return;
              }
            }

            otherParams[ky] = params[ky];

          });

          var rows = (function() {
            if (rowHeight) {
              const rows = {};
              rowHeight.map(data => (rows[data[0]] = { height: `${data[1]}px` }));
              return rows;
            }
            return {};
          })();

          otherParams.onload = this.onChange; // Hans addition

          otherParams.rows = rows;
          otherParams.tableOverflow = true;
          otherParams.onchange = this.onChange;
          otherParams.oninsertrow = this.onChange;
          otherParams.ondeleterow = this.onChange;
          otherParams.oninsertcolumn = this.onChange;
          otherParams.ondeletecolumn = this.onChange;
          otherParams.onsort = this.onChange;
          otherParams.onmoverow = this.onChange;
          otherParams.onchangeheader = this.onChangeHeader;

          if(showToolbar) {
            // Add toolbar to param
            otherParams.toolbar = [
              { type:'i', content:'undo', onclick:function() { excel.undo(); } },
              { type:'i', content:'redo', onclick:function() { excel.redo(); } },
              { type:'i', content:'save', onclick:function () { excel.download(); } },
              { type:'select', k:'font-family', v:['Arial','Verdana'] },
              { type:'select', k:'font-size', v:['9px','10px','11px','12px','13px','14px','15px','16px','17px','18px','19px','20px'] },
              { type:'i', content:'format_align_left', k:'text-align', v:'left' },
              { type:'i', content:'format_align_center', k:'text-align', v:'center' },
              { type:'i', content:'format_align_right', k:'text-align', v:'right' },
              { type:'i', content:'format_bold', k:'font-weight', v:'bold' },
              { type:'color', content:'format_color_text', k:'color' },
              { type:'color', content:'format_color_fill', k:'background-color' },
          ]
          }

          // If new instance of the table
          if(excel === null) {

            excel =  jexcel(container, otherParams);
            container.excel = excel;

            return;

          }

          var selection = excel.selectedCell;

          while (container.firstChild) {

            container.removeChild(container.firstChild);

          }

          excel = jexcel(container, otherParams);

          if(selection){

            excel.updateSelectionFromCoords(
                                            selection[0],
                                            selection[1],
                                            selection[2],
                                            selection[3]
                                            );

          }

          container.excel = excel;

        },

        resize: function(width, height) {

        },

        onChange: function(obj){

          if (HTMLWidgets.shinyMode) {

            var colType = this.columns.map(function(column){

              return column.type;

            })

            var colHeaders = this.colHeaders;

            Shiny.setInputValue(obj.id, {

                data:this.data,
                colHeaders: colHeaders,
                colType: colType,

              })
          }
        },

        onChangeHeader: function(obj, column, oldValue, newValue){

          if (HTMLWidgets.shinyMode) {

            var colHeaders = this.colHeaders;

            var newColHeader = colHeaders;
            newColHeader[parseInt(column)] = newValue;

            var colType = this.columns.map(function(column){

              return column.type;

            })

            Shiny.setInputValue(obj.id, {

                data:this.data,
                colHeaders: newColHeader,
                colType: colType

              })
          }
        }

      };
    }

  });


  if (HTMLWidgets.shinyMode) {

    // This function is used to set comments in the table
    Shiny.addCustomMessageHandler("excelR:setComments", function(message) {

      var el = document.getElementById(message[0]);
      if (el) {
        el.excel.setComments(message[1], message[2]);
      }
    });

    // This function is used to get comments  from table
    Shiny.addCustomMessageHandler("excelR:getComments", function(message) {

      var el = document.getElementById(message[0]);
      if (el) {

        var comments = message[1] ? el.excel.getComments(message[1]): el.excel.getComments(null);

        Shiny.setInputValue(message[0], {

           comments

        });

      }

    });
  }