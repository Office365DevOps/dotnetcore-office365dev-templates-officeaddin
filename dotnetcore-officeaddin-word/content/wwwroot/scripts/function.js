// 每次加载新页面时都必须运行初始化函数
(function () {
    Office.initialize = function (reason) {
        // 如果你需要初始化，可以在此进行。
    };
})();

function showSampleData() {
    console.log("entry");
    Word.run(function (ctx) {
        console.log("call the method");
        // 为文档正文创建代理对象。
        var body = ctx.document.body;

        // 将清空正文内容的命令插入队列。
        body.clear();
        // 将在 Word 文档正文结束位置插入文本的命令插入队列。
        body.insertText(
            "这是通过按钮代码直接插入的文本",
            Word.InsertLocation.end);


        Office.context.ui.displayDialogAsync("https://localhost:5001/dialog.html", { height: 30, width: 20, displayInIframe: true });

        // 运行排队的命令，并返回承诺表示任务完成
        return ctx.sync();
    });
}