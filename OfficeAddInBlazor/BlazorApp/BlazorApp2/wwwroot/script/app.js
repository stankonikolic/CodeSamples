function test(message) {
    Word.run(function (context) {

        context.document.body.insertText(message, Word.InsertLocation.start);
        return context.sync();
    })
}