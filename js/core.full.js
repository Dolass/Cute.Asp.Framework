var snda = {
	cookie: function(name, value, options){
	    if (typeof value != 'undefined') {
	        options = options || {};
	        if (value === null) {
	            value = '';
	            options = $.extend({}, options);
	            options.expires = -1;
	        }
	        var expires = '';
	        if (options.expires && (typeof options.expires == 'number' || options.expires.toUTCString)) {
	            var date;
	            if (typeof options.expires == 'number') {
	                date = new Date();
	                date.setTime(date.getTime() + (options.expires * 24 * 60 * 60 * 1000));
	            } else {
	                date = options.expires;
	            }
	            expires = '; expires=' + date.toUTCString(); // use expires attribute, max-age is not supported by IE
	        }
	        var path = options.path ? '; path=' + (options.path) : '';
	        var domain = options.domain ? '; domain=' + (options.domain) : '';
	        var secure = options.secure ? '; secure' : '';
	        document.cookie = [name, '=', encodeURIComponent(value), expires, path, domain, secure].join('');
	    } else {
	        var cookieValue = '';
	        if (document.cookie && document.cookie != '') {
	            var cookies = document.cookie.split(';');
	            for (var i = 0; i < cookies.length; i++) {
	                var cookie = $.trim(cookies[i]);
	                if (cookie.substring(0, name.length + 1) == (name + '=')) {
	                    cookieValue = decodeURIComponent(cookie.substring(name.length + 1));
	                    break;
	                }
	            }
	        }
	        return cookieValue;
	    }
	},

	checkAll:function (obj,elName) {
		if(obj.checked) {
			$(obj).closest("form").find("input[name="+elName+"]").attr("checked",true);
		}else {
			$(obj).closest("form").find("input[name="+elName+"]").attr("checked",false);
		}
	},
	comfirm:function (obj) {
		var obj = $(obj);
		if(obj.filter(":select").length > 0){
			var selectedElement=obj[0].options[obj[0].selectedIndex];
			var confirmValue=selectedElement.getAttribute("confirm")?selectedElement.getAttribute("confirm"):"";
			if(confirmValue!="") {
				if(!confirm(confirmValue)) {
					obj[0].selectedIndex=0;
				}
			}
			if(obj.selectedIndex>0) {
				$(obj).closest("form").submit();
				$(obj).attr("disabled","disabled");
			}
		}else{
			if(!confirm(obj.attr("confirm"))) {
				return false;
			}
		}
	},
	queryString: function(item) {
		var sValue = location.search.match(new RegExp("[\?\&]" + item + "=([^\&]*)(\&?)", "i"))
		sValue = sValue ? sValue[1] : sValue;
		return sValue == null ? "": sValue == undefined ? "" : decodeURIComponent(sValue);
	},
	hashString: function(item) {
		var sValue = location.hash.match(new RegExp("[\#\&]" + item + "=([^\&]*)(\&?)", "i"))
		sValue = sValue ? sValue[1] : sValue;
		return sValue == null ? "": sValue == undefined ? "" : decodeURIComponent(sValue);
	},
	
	displayPager: function(targerID, recordCount, pageSize, pageIndex, mark) {
		if ($(targerID).length==0) return;
		mark = !mark ? "&" : mark;
		var para = "";	
		if(para.toLowerCase().indexOf("pageid") == -1){
			if (para.indexOf("#") == -1) {
				para = location.href.replace(/\&?pageid\=[^&]+/gi,"");
				para += para.indexOf("?") == -1 ? "?" : "";
				para += mark + "pageid={0}";
				para = para.replace("?&", "?");
			}else {
				para = location.href.replace(/\#.+$/gi,"").replace(/pageid\=[^&]+/gi,"pageid={0}");
				para += mark + "&pageid={0}";
				para = para.replace("#&", "#");
			}
		}
		pageIndex = !pageIndex ? 1 : parseInt(pageIndex);
		var pagecount = Math.ceil(recordCount / pageSize);
		pagecount = Math.max(1, pagecount);
		var pagestr = "";
		var breakpage = 9;
		var currentposition = 4;
		var breakspace = 2;
		var maxspace = 4;
		var prevnum = pageIndex - currentposition;
		var nextnum = pageIndex + currentposition;
		if (prevnum < 1) prevnum = 1;
		if (nextnum > pagecount) nextnum = pagecount;
		pagestr += '<span class="pageIntroB">' + 
					'总数:<kbd class="p_total">' + recordCount + '</kbd>' + 
					' 每页:<kbd class="p_size">' + pageSize + '</kbd>' + 
					'</span>' + 
					'<div class="pageContorlB">' + (pageIndex == 1 ? '<a href="javascript:void(0)" class="p_disabled" disabled class="prev">前页</a>' : '<a href="' + para.replace("{0}", pageIndex - 1) + '" class="p_pre">前页</a>');
		if (prevnum - breakspace > maxspace) {
			for (i = 1; i <= breakspace; i++)
				pagestr += '<a href="' + para.replace("{0}", i) + '" class="p_page">' + i + '</a>';
			pagestr += '<span class="break">...</span>';
			for (i = pagecount - breakpage + 1; i < prevnum; i++)
				pagestr += '<a href="' + para.replace("{0}", i) + '" class="p_page">' + i + '</a>';
		} else {
		for (i = 1; i < prevnum; i++)
			pagestr += '<a href="' + para.replace("{0}", i) + '" class="p_page">' + i + '</a>';
		}
		for (i = prevnum; i <= nextnum; i++) {
			pagestr += (pageIndex == i) ? '<strong class="p_cur">' + i + '</strong>' : '<a href="' + para.replace("{0}", i) + '" class="p_page">' + i + '</a>';
		}
		if (pagecount - breakspace - nextnum + 1 > maxspace) {
			for (i = nextnum + 1; i <= breakpage; i++)
				pagestr += '<a href="' + para.replace("{0}", i) + '" class="p_page">' + i + '</a>';
			pagestr += '<span class="break">...</span>';
			for (i = pagecount - breakspace + 1; i <= pagecount; i++)
				pagestr += '<a href="' + para.replace("{0}", i) + '" class="p_page">' + i + '</a>';
		} else {
			for (i = nextnum + 1; i <= pagecount; i++)
				pagestr += '<a href="' + para.replace("{0}", i) + '" class="p_page">' + i + '</a>';
		}
		pagestr += (pageIndex == pagecount) ? '<a class="p_disabled" href="javascript:void(0)" disabled>后页 </a></div>' : '<a href="' + para.replace("{0}", pageIndex + 1) + '" class="p_next">后页</a></div>';
		if (pagecount <= 1) {
			$(targerID).html("");
		}
		else {
			$(targerID).html(pagestr);
		}
	}

}