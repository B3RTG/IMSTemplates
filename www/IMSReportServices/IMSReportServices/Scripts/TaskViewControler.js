/*aux format date*/
Date.prototype.ddmmyyy_hhMMss = (function (date_separator, time_separator) {
    datetime = this;
    date_separator = (date_separator ? date_separator : "-");
    time_separator = (time_separator ? time_separator : ":");

    var yyyy = datetime.getFullYear().toString();
    var mm = (datetime.getMonth() + 1).toString(); // getMonth() is zero-based
    var dd = datetime.getDate().toString();
    var hh = datetime.getHours().toString();
    var min = datetime.getMinutes().toString();
    var ss = datetime.getSeconds().toString();

    return yyyy + date_separator + (mm[1] ? mm : "0" + mm[0]) + date_separator + (dd[1] ? dd : "0" + dd[0]) + " " +
           (hh[1] ? hh : "0" + hh[0]) + time_separator + (min[1] ? min : "0" + min[0]) + time_separator + (ss[1] ? ss : "0" + ss[0]); // padding
});

var app_name = 'IMSReportService';
function TaskViewControler(tableID) {
    var _self = this;
    _self.HTMLTable = $('#'+tableID);
    _self.oDataTable = null;
    _self.Tasks = null;

    _self.JobsUri = '/'+ app_name + '/api/Job/';
    _self.TaskUri = '/'+ app_name + '/api/Task/';

    _self.ListAllTask = function () {
        $.ajax({
            type: 'GET',
            url: _self.TaskUri,
            dataType: 'json',
            contentType: 'application/json'
        }).success(function (data) {
            _self.Tasks = data;
            //console.log(_self.Tasks)
            //for reload
            if (_self.oDataTable != null) {
                _self.oDataTable.destroy();
            }
            /*
            <th>ID</th>
                    <th>Job Code</th>
                    <th>Tittle</th>
                    <th>Create Date</th>
                    <th>End Date</th>
                    <th>Status</th>
                    <th>Get Result</th>
            */
            //setup datatable
            _self.oDataTable = _self.HTMLTable.DataTable({
                data: data,
                order: [[3,"desc"]],
                columns: [
                    { data: 'TaskID', visible:false },
                    { data: 'oJob.JOBCODE', visible: false },
                    { data: 'oJob.Tittle'},
                    { data: 'CreateDate'},
                    { data: 'UpdateDate' },
                    { data: 'StatusCurrent' },
                    { data: 'StatusFinal', className: "dt-center", orderable: false },
                ],
                columnDefs: [
                    {
                        render: function (data, type, row) {
                            //var btnThrow = $("<button></button>", { class: "btn btn-default btn-sm", "data-job-id": row.JOBID });
                            if (data == "DONE")//&& row.JOBID != 4
                                //return "<button href='/Home/CreateTask?JobID=" + row.JOBID + "' class='btn btn-default btn-sm create' data-job-id='" + row.JOBID + "' data-job-code='" + row.JOBCODE + "'>Lanzar</button>";
                                //return 'descargar';
                                return "<a href='../TaskManager.aspx?requestType=getFile&TaskID=" + row.TaskID + "'><img src='../images/excel.png' /></a>";
                                //btnThrow.text = "Lanzar";
                            else
                                return '';
                                //return "<button href='/Home/CreateTask?JobID=" + row.JOBID + "' class='btn btn-default btn-sm disabled' data-job-id='" + row.JOBID + "' data-job-code='" + row.JOBCODE + "'>Running</button>";
                            //btnThrow.text = "Ver estado";

                            return btnThrow;
                        },
//                        orderable:false,
                        targets: 6
                    },
                    {
                        render: function (data, type, row) {
                            var d = new Date(data);
                            //console.log(d.getFullYear());
                            //return ddmmyyy_hhMMss(d);
                            //return data;
                            return (data ? d.ddmmyyy_hhMMss() : "");
                        },
                        targets: [3,4]
                    },
                    {
                        targets: 5,
                        render: function (data, type, row) {
                            var sResult = "";
                            switch (data) {
                                case "TODO":
                                    sResult = "En cola";
                                    break;
                                case "IMDO":
                                    sResult = "Importado. Falta Formateado.";
                                    break;
                                case "DONE":
                                    sResult = "Finalizado.";
                                    break;
                                case "ERRO":
                                    sResult = "Error";
                                    break;
                                default:
                                    sResult = data;
                            }
                            return sResult;
                        }

                    }
                ]
            });
        }).fail(function (jqXHR, textStatus, errorThrown) {
            //_self.error(errorThrown);
            console.log("Error");
        });
    }

    // Fetch the initial data.
    //getAllJobs();
};