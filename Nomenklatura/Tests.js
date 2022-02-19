function consLogIDE(msg, vsCode) {
    // в зависимости от IDE делавть выввод
    if (vsCode) {
        console.log(msg);
    } else {
        Browser.log(msg);
    }
}

consLogIDE();
