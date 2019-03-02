// 每次加载新页面时都必须运行初始化函数
(function () {
    Office.initialize = function (reason) {
        // 如果你需要初始化，可以在此进行。
    };
})();


function showSampleData() {
    Excel.run(function (ctx) {
        var sheet = ctx.workbook.worksheets.getActiveWorksheet();
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];
        // 将向电子表格写入示例数据的命令插入队列
        sheet.getRange("B3:D5").values = values;

        Office.context.ui.displayDialogAsync("https://localhost:5001/dialog.html", { height: 30, width: 20, displayInIframe: true });

        // 运行排队的命令，并返回承诺表示任务完成
        return ctx.sync();
    });
}