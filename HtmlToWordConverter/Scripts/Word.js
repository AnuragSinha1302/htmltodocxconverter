$('#btnGenerateWord').on('click', () => {
    $.ajax(
        {
            url: '/Word/GenerateWord',
            type: 'GET',
            success: function (response) {
                console.log(response);
            },
            error: function (error) {
                console.log(error);
            }
        });
});