var app_name = '';
function JobViewModel() {
    var _self = this;
    _self.Jobs = null;
    //self.Jobs = ko.observableArray();
    //self.error = ko.observable();
    _self.JobsUri = app_name + 'api/Job/';
    _self.TaskUri = app_name + 'api/Task/';
    _self.JobListTable = null;


    _self.ListAllJobs = function () {
        $.ajax({
            type: 'GET',
            url: _self.JobsUri,
            dataType: 'json',
            contentType: 'application/json'
        }).success(function (data) {
            _self.Jobs = data;
            if (_self.JobListTable != null)
            {
                _self.JobListTable.destroy();
            }

            _self.JobListTable = $('#tblJobList').DataTable({
                data: data,
                paging: false,
                searching: false,
                order: [[2, "desc"]],
                columns: [
                    { data: 'CurrentTaskStatus', className: 'SentButton', defaultContent: '', orderable: false, width: "10%" },
                    { data: 'JOBID', name: "Job ID", width: "10%", visible: false },
                    { data: 'JOBCODE', name: "Job Code", width: "20%", visible: false },
                    { data: 'Tittle', name: "Report Tittle" }
                ],
                columnDefs: [
                    {
                        render: function (data, type, row) {
                            //var btnThrow = $("<button></button>", { class: "btn btn-default btn-sm", "data-job-id": row.JOBID });
                            if (data == "FREE" )//&& row.JOBID != 4
                                return "<button href='/Home/CreateTask?JobID=" + row.JOBID + "' class='btn btn-default btn-sm create' data-job-id='" + row.JOBID + "' data-job-code='" + row.JOBCODE + "'>Lanzar</button>";
                                //btnThrow.text = "Lanzar";
                            else
                                return "<button href='/Home/CreateTask?JobID=" + row.JOBID + "' class='btn btn-default btn-sm disabled' data-job-id='" + row.JOBID + "' data-job-code='" + row.JOBCODE + "'>Running</button>";
                                //btnThrow.text = "Ver estado";

                            return btnThrow;
                        },
                        orderable: false,
                        targets: 0
                    }
                ]
            });

            $('.btn.create').on("click", function () {
                var bCurrentButton = $(this);
                var oTask = new TaskViewMode(bCurrentButton.data('job-id'), bCurrentButton.data('job-code'));

                oTask.loadJobData();
                
            });

        }).fail(function (jqXHR, textStatus, errorThrown) {
            //_self.error(errorThrown);
            console.log("Error");
        });
    }

    // Fetch the initial data.
    //getAllJobs();
};


function TaskViewMode(iJobID, sJobCode) {
    var _self = this;
    _self.JobID = iJobID;
    _self.JobCode = sJobCode;
    _self.JobsUri = app_name + 'api/Job/';
    _self.TaskUri = app_name + 'api/Task/';

    var divForm = "#task_form";

    _self.loadJobData = function () {
        //load job information
        $.ajax({
            type: 'GET',
            url: _self.JobsUri +'/' + iJobID,
            dataType: 'json',
            contentType: 'application/json'
        }).success(function (data) {
            loadJob(data);
        });

        
    }

    function loadJob(job)
    {
        $("#JobCODE").val(job.JOBCODE);
        $("#JobID").val(job.JOBID);
        var dContainer = $(divForm);
        var dFiles = dContainer.find(".files");
        dFiles.find("input:file").remove();

        for(oFile in job.InputParameters.Files)
        {
            var current_file = job.InputParameters.Files[oFile];
            var file_control = $('<input>', { type: "file", name: current_file.Name, id: current_file.Name , class:"form-control", accept:".xlsx,.xls"});
            file_control.appendTo(dFiles);
        }

        dContainer.show();
    }

}

function FileUploadControl(form_name, CurrentJobView) {
    _self = this;
    _self.oForm = $("#" + form_name);
    _self.divForm = $("#task_form");
    _self.oJobViewModel = CurrentJobView;

    $('#bCancel').on("click", function () {
        _self.divForm.hide();
    });
    _self.oForm.submit(function (event) {
        event.preventDefault();
        $("#loading").show();
        var url = 'TaskManager.aspx?requestType=setTask&iframe=true';

        //console.log(_self);
        //prepare data to sent
        var aData = [];
        aData.push({ "name": "jobID", "value": $("#JobID").val() });
        aData.push({ "name": "jobCODE", "value": $("#JobCODE").val() });

        //prepare files
        var aFiles = $(":file", _self.oForm);

        /*envio de formulario*/
        $.ajax(url, {
            data: aData,
            files: aFiles,
            iframe: true,
            processData: false
        }).complete(function (data) {
            var ResponseObject = data.responseJSON;//$.parseJSON(data);
            console.log(ResponseObject);
            var divAlert = $('#TaskMessage');
            var pMessage = divAlert.find('p')
            pMessage.text(ResponseObject.message);
            _self.divForm.hide();
            

            if (ResponseObject.Status == 'ERROR')
            {//fail
                if (divAlert.hasClass("alert-success")) {
                    divAlert.toggleClass("alert-success");
                    divAlert.toggleClass("alert-danger");
                }
            } else { //OK
                if (!divAlert.hasClass("alert-success")) {
                    divAlert.toggleClass("alert-success");
                    divAlert.toggleClass("alert-danger");
                }
            }
            divAlert.show();
            _self.oJobViewModel.ListAllJobs();
            $("#loading").hide();
            /*
            if (result.Status == 0) {
                if (result.Data.Message != null) result.Message = result.Data.Message;
                else result.Message = "Error cargando fichero.";
            }

            if (result.Done) {
                showMessage('alert', result.Message, 'success', null);
            } else {
                showMessage('alert', result.Message, 'error', null);
            }*/

            /*
            disableFormControl('submit', 'enable');
            $('#resolution').prop('disabled', false);
            $('#resolution').selectpicker('refresh');
            */
        }).fail(function (jqXHR, status, errorThrown) {
            console.log(errorThrown);
            console.log(jqXHR.responseText);
            console.log(jqXHR.status);
            /*
            disableFormControl('submit', 'enable');
            $('#resolution').prop('disabled', false);
            $('#resolution').selectpicker('refresh');
            */
        })
        
        return false;
    });
    
}


//ko.applyBindings(new ViewModel());