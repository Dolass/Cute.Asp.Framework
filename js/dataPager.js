function dataPager(fn) {
    var _self = this;
    var _executeSource = fn;
    
    this.recordCount = 0;
    this.pageCount = 0;
    this.pageIndex = -1;
    this.pageSize = 20;
    this.groupCount = 0;
    this.groupIndex = -1;
    this.groupSize = 10;
    this.sortFields = [];
    this.searchFields = [];
    
    this.setRecordCount = function(count) {        
        if(isNaN(count) || count < 0) return; 
        _self.recordCount = count;        
        if(_self.pageSize > 0) { var num = Math.floor(count / _self.pageSize); _self.pageCount = count % _self.pageSize == 0 ? num : (num + 1); }
        else _self.pageCount = 0; 
        if (_self.groupSize > 0) { var num = Math.floor(_self.pageCount / _self.groupSize); _self.groupCount = _self.pageCount % _self.groupSize == 0 ? num : (num + 1); }
        else _self.groupCount = 0;
    }
    
    this.firstPage = function() {
        _self.pageIndex = 0;
        _self.groupIndex = 0;
        _executeSource(_self);
    }
    
    this.priviousPage = function() {
        if(_self.pageIndex <= 0) {
            _self.firstPage();
        }
        else {
            _self.pageIndex--;
            var firstIndex = _self.groupIndex * _self.groupSize;
            if (_self.pageIndex < firstIndex && _self.groupIndex > 0) _self.groupIndex--;
            _executeSource(_self);
        }
    }
    
    this.lastPage = function() {
        _self.pageIndex = _self.pageCount > 0 ? _self.pageCount - 1 : 0;
        _self.groupIndex = _self.groupCount > 0 ? _self.groupCount - 1 : 0;
        _executeSource(_self);
    }
    
    this.nextPage = function() {
        if (_self.pageIndex >= _self.pageCount - 1) {
            _self.lastPage();
        }
        else {
            _self.pageIndex++;
            var lastIndex = (_self.groupIndex + 1) * _self.groupSize - 1;
            if (_self.pageIndex > lastIndex && _self.groupIndex < _self.groupCount - 2) _self.groupIndex++;
            _executeSource(_self);
        }
    }   
        
    this.gotoPage = function(index) {
        if(isNaN(index) || _self.pageIndex == index) {
            _self.currentPage();
        }
        else {
            if(index <= 0) {
                _self.firstPage();
            }
            else if (index >= _self.pageCount - 1) {
                _self.lastPage();
            }
            else {
                _self.pageIndex = index;
                var num = Math.floor(_self.pageIndex / _self.groupSize);
                _self.groupIndex = _self.pageIndex % _self.groupSize == 0 ? num : (num + 1);
                _executeSource(_self);
            }
        }
    }
    
    this.currentPage = function() {
        _executeSource(_self);
    }
    
    this.firstGroup = function () {
        if (_self.pageIndex < 0 || _self.pageIndex > _self.groupSize - 1) _self.gotoPage(0);
        else _executeSource(_self);
    }
    
    this.priviousGroup = function() {
        if (_self.groupIndex <= 0) {
            _self.firstGroup();
        }
        else {
            var firstIndex = (_self.groupIndex - 1) * _self.groupSize;
            var lastIndex = _self.groupIndex * _self.groupSize - 1;
            if (_self.pageIndex < firstIndex || _self.pageIndex > lastIndex) _self.gotoPage(firstIndex);
            else _executeSource(_self);
        }
    }
    
    this.lastGroup = function() {
        var firstIndex = (_self.groupCount - 1) * _self.groupSize;
        var lastIndex = _self.groupCount * _self.groupSize - 1;
        if (_self.pageIndex < firstIndex || _self.pageIndex > lastIndex) _self.gotoPage(firstIndex);
        else _executeSource(_self);
    }
    
    this.nextGroup = function () {
        if (_self.groupIndex >= _self.groupCount - 1){
            _self.lastGroup();
        }
        else {
            var firstIndex = (_self.groupIndex + 1) * _self.groupSize;
            var lastIndex = (_self.groupIndex + 2) * _self.groupSize - 1;
            if (_self.pageIndex < firstIndex || _self.pageIndex > lastIndex) _self.gotoPage(firstIndex);
            else _executeSource(_self);
        }
    }
    
    this.searchPage = function(fields, searchResult) {
        if(!searchResult) _self.searchFields = [];
        if(fields) for(var i=0; i<fields.length; i++) { _self.searchFields.push(fields[i]); };
        _self.setRecordCount(0);
        _self.sortFields = []; //todo:remove this line?
        _self.firstPage();
    }
    
    this.addSearch = function(name, value) {
        _self.searchFields.push({name:name, value:value});
        _self.setRecordCount(0);
    }
    
    this.sortPage = function(name) {
        if(!name) { _self.firstPage(); return; }
        if(_self.sortFields == []) { 
            _self.sortFields.push({name:name, isAsc:1}); 
        }
        else {
            if(_self.sortFields[0].name == name) _self.sortFields[0].isAsc = !_self.sortFields[0].isAsc;
            else _self.sortFields[0] = {name:name, isAsc:1};
        }
        _self.firstPage();
    }
    
    this.addSort = function(name, isAsc) {
        _self.sortFields.push({name:name, isAsc:isAsc?1:0});
    }
    
    this.buildParams = function() {
        var p = [];
        p.push({name:"recordCount", value:_self.recordCount});
        p.push({name:"pageIndex", value:_self.pageIndex});
        p.push({name:"groupIndex", value:_self.groupIndex});        
        var a = [];
        for(var i=0; i<_self.searchFields.length; i++) {
            a.push("s_"+_self.searchFields[i].name);
            p.push({name:"s_"+_self.searchFields[i].name, value:_self.searchFields[i].value});
        }
        p.push({name:"searchFields", value:a.join(",")});
        a = [];
        for(var i=0; i<_self.sortFields.length; i++) {
            a.push("o_"+_self.sortFields[i].name);
            p.push({name:"o_"+_self.sortFields[i].name, value:_self.sortFields[i].isAsc ? 1 : 0});
        }
        p.push({name:"sortFields", value:a.join(",")});
        for(var i=0; i<arguments.length; i++) {
            p.push(arguments[i]);
        }
        return p;
    }
}