// inline ad randomiser for Ad Insert plugin for Wordpress.

// declare container array
var ads = [];
var imgID = "inlinead";

// populate array with wordpress image locations taken from WP media manager
ads.push ("https://domain.com/path/to/image/inline_2.png");
ads.push ("https://domain.com/path/to/image/inline_3.png");
ads.push ("https://domain.com/path/to/image/inline_4.png");
ads.push ("https://domain.com/path/to/image/inline_5.png");
ads.push ("https://domain.com/path/to/image/inline_6.png");
ads.push ("https://domain.com/path/to/image/inline_7.png");

// get number of ads in array
num_ads = ads.length;

// get random ad subscript
random_pick = Math.floor(Math.random() * (num_ads + 1));

// get random ad URL
random_ad_url = ads [random_pick];

// replace src based on IMG id
document.getElementById(imgID).src=random_ad_url;
