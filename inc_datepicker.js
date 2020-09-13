<!--  
	function openWin( windowURL, windowName, windowFeatures ) 
	{ 
		window.open( windowURL, windowName, windowFeatures ); 
	}

	function makeRemote(url) 
	{
		remote = window.open(url,"remotewin","width=400,height=500,scrollbars=1");
		remote.location.href = url;
		if (remote.opener == null) remote.opener = window;
	}

	function makeRemoteEx( i_strUrl, i_lngWidth, i_lngHeight, i_blnScrollBars) 
	{
		var strScrollBars;
		if (i_blnScrollBars)
			strScrollBars = "1";
		else
			strScrollBars = "0";
			
		remote = window.open(i_strUrl,"remotewin","width=" + i_lngWidth + ",height=" + i_lngHeight + ",scrollbars=" + strScrollBars);
		remote.location.href = i_strUrl;
		if (remote.opener == null) remote.opener = window;
	}

	function _ShowPopupCalendar( i_strFormName, i_strFieldName, i_blnHistory )
	{
		var strURL = 'pop_datepicker.asp?FormName=' + i_strFormName + '&FieldName=' + i_strFieldName;
		
		if (i_blnHistory)
			strURL = strURL + '&History=on';
		else
			strURL = strURL + '&History=off';
			
		window.datefield = document.forms[i_strFormName][i_strFieldName];
		var strDate = window.datefield.value;
		window.open( strURL + '&date=' + strDate, 'cal', 'width=300,height=340' );
	}      

        function Mid(str, start, len)
        {
                if (start < 0 || len < 0) return "";

                var iEnd, iLen = String(str).length;
                if (start + len > iLen)
                        iEnd = iLen;
                else
                        iEnd = start + len;

                return String(str).substring(start,iEnd);
        }
//-->