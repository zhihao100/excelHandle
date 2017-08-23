ExcelHandle.service('excelAlert', ['$modal', '$http', function ($modal, $http) {
    /*消息提示模式窗口*/
    this.show = function (config) {
        var DEFAULT = {
            templateUrl: 'views/common/popupMessage.html',
            controller: function ($scope, $modalInstance, items) {
                $scope.results = items;
                // 确认按钮
                $scope.ok = function () {
                    $modalInstance.close();
                };
                // 取消按钮
                $scope.cancel = function () {
                    $modalInstance.dismiss('cancel');
                };
            },
            size: 'sm',
            title: '提示消息',
            msg: ''
        };
        var cfg = (Object.prototype.toString.call(config) === "[object String]") ? $.extend(DEFAULT, {msg: config}) : $.extend(DEFAULT, config);

        //创建弹框对象
        var modalInstance = $modal.open({
            templateUrl: cfg.templateUrl,
            controller: cfg.controller,
            size: cfg.size,
            resolve: {
                items: function () {
                    return {title: cfg.title, msg: cfg.msg};
                }
            }
        });
        return modalInstance;
    };
}]);