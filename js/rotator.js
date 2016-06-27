 $(window).load(function() { //start after HTML, images have loaded

    var InfiniteRotator = {
        init: function() {
            //initial fade-in time (in milliseconds)
            var initialFadeIn = 1000;

            //interval between items (in milliseconds)
			// 0.3 minutes = 30000
			// 2.5 minutes = 150000
			// 5.0 minutes = 300000
            var itemInterval = 50000;

            //cross-fade time (in milliseconds)
            // var fadeTime = 5000; set in code

            //count number of items
            var numberOfItems = $('.rotating-item').length;

            //set current item
            var currentItem = 0;

            //show first item
            $('.rotating-item').eq(currentItem).fadeIn(initialFadeIn);

            //loop through the items		
            var infiniteLoop = setInterval(function() {
                $('.rotating-item').eq(currentItem).fadeOut(fadeTime);

                if (currentItem == numberOfItems - 1) {
                    currentItem = 0;
                } else {
                    currentItem++;
                }
                $('.rotating-item').eq(currentItem).fadeIn(fadeTime);

            }, itemInterval);
        }
    };

    InfiniteRotator.init();

});