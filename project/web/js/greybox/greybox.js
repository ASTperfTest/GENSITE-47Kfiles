GB_myShow = function(caption, url, /* optional */ height, width, callback_fn) {
    var options = {
        type: 'page',
        caption: caption,
        height: height || 500,
        width: width || 800,
        center_win: true, 
        fullscreen: false,
        show_loading: false,
        callback_fn: callback_fn
    }
    var win = new GB_Window(options);
    return win.show(url);
}


