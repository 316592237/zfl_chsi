function imgSwipe(){
    //图片可预览效果的实现。
    $('.materialboxed').on('click',function(event) {
        var imgArray = [],imgIndex=0,$this = $(this);
        var curImageSrc = $this.children('img').eq(0).attr('src');
        if (curImageSrc && !$this.attr('href')) {
            $('.materialboxed').each(function(index, el) {
                var $img = $(this).children('img').eq(0),
                    itemSrc = $img.attr('src');
                if(itemSrc == curImageSrc){
                    imgIndex = index;
                }    
                if(itemSrc.indexOf('thumbnail')>=0){
                    itemSrc = itemSrc.replace('thumbnail','originalimage');
                }
                var bodyWidth = $(window.document.body).width(),
                    width = $img.width(),
                    height = $img.height();
                var h = (bodyWidth/width)*height;
                imgArray.push({src:itemSrc,w:bodyWidth,h:h});
                
            });
            openPhotoSwipe(imgIndex,imgArray);
        }
    });
}
//调用PhotoSwiper方法，图片放大效果
function openPhotoSwipe(index,imgArr){
    var pswpElement = document.querySelectorAll('.pswp')[0];
    var $content = $(pswpElement).find('.pswp__container');
    var items =  $content.find('.pswp__item').length;
    var imgItems = imgArr.length;
    var n = imgItems - items;
    if(n >0){
        //动态设置item，否则报错，item和图片个数相同
        for(var i =0 ;i<n ;i++){
            $('<div class="pswp__item"></div>').appendTo($content);
        }
    }
    var options = {
        history: false,
        focus: false,
        showAnimationDuration: 0,
        hideAnimationDuration: 0,
        index:index
    };
    var gallery = new PhotoSwipe(pswpElement, PhotoSwipeUI_Default, imgArr, options);
    gallery.init();
}


/**
 * description：标题不恰当换行
 * @param cName: 目标元素的cName
 * @author ：huanghui
 */
function changeLine (cName) {
    var $title = $("."+cName);
    var width = window.innerWidth,
        height = $title.height(),
        html   = $title.html();
    //高度大于30肯定已换行，最小width为320
    if(height>30 && width>=320){
        var index = html.lastIndexOf('认证报告');
        //字数较多时，避免换三行
        if(width<=360 && index>=15){
            index = html.lastIndexOf('成绩认证报告');
        }
        var str = html.substring(0,index)+'<br/>'+html.substring(index);
        $title.html(str);
    }
}