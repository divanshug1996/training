$(document).ready(
			function() {
				$('.process-spreadsheet .file-upload input').change(
						function() {
							$('.process-spreadsheet .file-upload p').text(
									this.files.length + " file(s) selected");
						});
			});

	function validate(file) {
		var ext = file.split(".");
		ext = ext[ext.length - 1].toLowerCase();
		var arrayExtensions = [ "xls", "xlsx", "xlsm" ];
		$('#uploadFileBtn').removeAttr("disabled");

		if (arrayExtensions.lastIndexOf(ext) == -1) {
			alert("Wrong file type. Upload only files with extensions .xls, .xlsx or .xlsm");
			$("#fileToUpload").val("");
			$('#uploadFileBtn').attr('disabled', 'disabled');
		}
	}