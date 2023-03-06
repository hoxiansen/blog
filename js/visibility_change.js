var originTitle = document.title;
var titleTime;
document.addEventListener('visibilitychange', function () {
    if (document.hidden) {
        document.title = '(*σ´∀`)σ 我走噜~ ' + originTitle;
        if (titleTime != null) {
            clearTimeout(titleTime);
        }
    } else {
        document.title = '(*´∇｀*) 我来力~ ' + originTitle;
        titleTime = setTimeout(function () {
            document.title = originTitle;
        }, 2000);
    }
});