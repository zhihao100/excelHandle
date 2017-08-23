var ExcelHandle = angular.module("ExcelHandle", [ 'ui.bootstrap',
		'angularFileUpload' ]);
ExcelHandle.controller("ExcelHandleController",
		[ "$http", "$scope", "$modal", function($http, $scope, $modal) {
			$scope.importExcel = function() {
				var params = {
					url : '/getImportTemplateExcel'
				};
				$modal.open({
					templateUrl : 'views/importExcelModal.html',
					controller : 'importExcelModalController',
					size : 'lg',
					resolve : {
						params : function() {
							return params;
						}
					}
				});
			}
		} ]).controller(
		"importExcelModalController",
		[
				"$http",
				"$scope",
				"params",
				"$modal",
				"$modalInstance",
				"FileUploader",
				"excelAlert",
				function($http, $scope, params, $modal, $modalInstance,
						FileUploader, excelAlert) {
					$scope.vm = {};
					$scope.vm.url = params.url;
					// 导入
					var uploader = $scope.uploader = new FileUploader({
						url : "/importExcel",
						queueLimit : 1,
						removeAfterUpload : true
					});
					$scope.clearItems = function() {
						// 重新选择文件时，清空队列，达到覆盖文件的效果
						uploader.clearQueue();
					}
					uploader.onAfterAddingFile = function(fileItem) {
						$scope.fileItem = fileItem._file;
					}
					// 保存
					$scope.ok = function() {
						uploader.uploadAll();
						uploader.onSuccessItem = function(fileItem, response,
								status, headers) {
							if (response != "") {
								excelAlert.show(response);
								$modalInstance.close('success');
							} else {
								excelAlert.show("导入成功!");
								$modalInstance.close('success');
							}
						};
					}
					// 取消按钮
					$scope.cancel = function() {
						// 跳转到列表页面
						$modalInstance.dismiss('cancel');
					};
				} ]);
