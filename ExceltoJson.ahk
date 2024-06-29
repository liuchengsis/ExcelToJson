; 从xlsx_path目录检索xlsx文件并导出json至json_path目录,exportmode决定导出模式
excel_tojson(xlsx_path,json_path,exportmode:=0)
{
    static type_True:= Object("json_value", "true", "value", 1)
        , type_False:= Object("json_value", "false","value", 0)
        , type_Null := Object("json_value", "null", "value","")
    dest:=json_path, xlsxlist:=[] ;要导出的文件列表
    loop, Files, %xlsx_path%\*.xlsx, R
    {
        if A_LoopFileAttrib contains H,S
            continue
        xlsxlist.Push(A_LoopFileDir . "\" . A_LoopFileName)
    }
    xl := ComObjCreate("Excel.Application"), xl.DisplayAlerts:=false, win_id:=xl.hWnd
    WinGet, p_pid, PID, ahk_id %win_id%
    key_main_list:="0", sheetname:="", parent_name:="", master_table:={}, follower_table_haskey:={}, follower_table_nokey:={}
    for k, v in xlsxlist
    {
        wb := xl.workbooks.open(v)
        loop % wb.worksheets.Count
        {
            ws:= wb.worksheets[A_Index], ws.Activate
            if SubStr(ws.Name,1,1)="!"
                continue
            ahk_obj:={}, table_type:=1, table_struct:={}, row:=1, col:=1, index_title:=0, owner_col:=0,text:="", while_ok:=1, blank:=0 ;table_struct[第几列]=>[字段名,类型码]
                , InStr(ws.Name,"@") ? (table_type++, parent_name:=SubStr(ws.Name, InStr(ws.Name,"@")+1), sheetname:=SubStr(ws.Name, 1, InStr(ws.Name,"@")-1)) : sheetname:=ws.Name
            while while_ok
            {
                text:=ws.cells[row,col].value
                if(!index_title) ;先找标头行
                {
                    text ? blank:=0 : ++blank
                    if(blank>=10)
                        Break
                    pos_sharp:=InStr(text, "#"), (pos_sharp && SubStr(text, pos_sharp+1)="id") ? index_title:=row : row++
                    continue
                }
                if(row=index_title && text) ;读标头，确定表的结构
                {
                    if(SubStr(text,1,1)!="!")
                        pos_sharp:=InStr(text, "#"), table_struct.Push([SubStr(text, 1, pos_sharp-1),SubStr(text, pos_sharp+1)]), ((table_type>=2 && RegExMatch(SubStr(text, pos_sharp+1),"i)^owner$")) && owner_col:=col)
                    col++
                    continue
                }
                else if(row=index_title && !text)
                {
                    row++, col:=1
                    continue
                }
                if(row>index_title) ;读内容
                {
                    table_type>=2 ? ctrl_value:=ws.cells[row,owner_col].text : ctrl_value:=ws.cells[row,1].text
                    if(ctrl_value="")
                    {
                        while_ok:=0
                        Break
                    }
                    key_main:=ws.cells[row,1].text
                    if(key_main!="")
                    {
                        if key_main not in %key_main_list% ; 主键重名检测
                            key_main_list.=","key_main, ahk_obj[key_main]:={}, (table_type==2 && table_type++)
                        else
                        {
                            row++
                            continue
                        }
                    }
                    else
                        row_obj:={}
                    while col<=table_struct.Count()
                    {
                        text:=ws.cells[row,col].value, (text="null" && text:=type_Null.Clone())
                        Switch table_struct[col][2]
                        {
                        Case "{}", "[]":
                            text:=json_toobj(text).Clone()
                                , key_main!="" ? ahk_obj[key_main][(table_struct[col][1])]:=text : row_obj[(table_struct[col][1])]:=text
                        Case "bool", "boolean":
                            text:=ws.cells[row,col].text
                            if(RegExMatch(text,"i)^true$"))
                                text:=type_True.Clone()
                            else if(RegExMatch(text,"i)^false$"))
                                text:=type_False.Clone()
                            else
                                errlog("Type_Error_Boolean",wb.Name,ws.Name,row,col)
                            key_main!="" ? ahk_obj[key_main][(table_struct[col][1])]:=text : row_obj[(table_struct[col][1])]:=text
                        Case "owner":
                            key_main!="" ? ahk_obj[key_main]["owner"]:=text : row_obj["owner"]:=text
                        Case "id":
                            (key_main!="" && ahk_obj[key_main][(table_struct[col][2])]:=text)
                        Case "timestamp":
                            txet:=Unix_timestamp(text), key_main!="" ? ahk_obj[key_main][(table_struct[col][1])]:=text : row_obj[(table_struct[col][1])]:=text
                        Case "timestamp13":
                            txet:=Unix_timestamp(text,13), key_main!="" ? ahk_obj[key_main][(table_struct[col][1])]:=text : row_obj[(table_struct[col][1])]:=text
                        Default:
                            key_main!="" ? ahk_obj[key_main][(table_struct[col][1])]:=text : row_obj[(table_struct[col][1])]:=text
                        }
                        col++
                    }
                    (key_main="" && ahk_obj.Push(row_obj.Clone()))
                    if(!while_ok)
                        Break
                }
                row++, col:=1
            }
            Switch table_type ;将ahk对象缓存
            {
            Case 1:
                if(!master_table.HasKey(sheetname))
                    master_table[sheetname]:=ahk_obj.Clone()
            Case 2:
                if(!follower_table_nokey.HasKey(sheetname))
                    follower_table_nokey[sheetname]:=[ahk_obj.Clone(),parent_name,0]
            Case 3:
                if(!follower_table_haskey.HasKey(sheetname))
                    follower_table_haskey[sheetname]:=[ahk_obj.Clone(),parent_name,0]
            }
            ws:=""
        }
        wb.Close(), wb:=""
    }
    ; 先备份缓存，以便依据不同模式导出
    master_table_cache:={}, follower_table_nokey_cache:={}, follower_table_haskey_cache:={}
    for k, v in master_table
        master_table_cache[k]:=ahk_tojson(v)
    for k, v in follower_table_nokey
        follower_table_nokey_cache[k]:=ahk_tojson(v[1])
    for k, v in follower_table_haskey
        follower_table_haskey_cache[k]:=ahk_tojson(v[1])
    ; 从缓存中取出对象，依据引用关系重组对象后导出json
    for k, v in follower_table_nokey ;k->sheetname, v->[ahk_obj.Clone(),parent_name,0]
    {
        if(v[3]==0)
        {
            if(follower_table_haskey.HasKey(v[2])) ;先在带主键的从表集合中查找是否存在被引用的主表
            {
                for i, j in v[1] ;i->数组下标 j->row_obj
                {
                    parent_key:=j.owner, j.Delete("owner")
                    if(!follower_table_haskey[v[2]][1][parent_key].HasKey(k))
                        follower_table_haskey[v[2]][1][parent_key][k]:=j
                    else if(follower_table_haskey[v[2]][1][parent_key].HasKey(k) && typeof(follower_table_haskey[v[2]][1][parent_key][k])=="object")
                        obj_cache:=follower_table_haskey[v[2]][1][parent_key][k].Clone(), follower_table_haskey[v[2]][1][parent_key][k]:=[], follower_table_haskey[v[2]][1][parent_key][k].Push(obj_cache,j)
                    else if(follower_table_haskey[v[2]][1][parent_key].HasKey(k) && typeof(follower_table_haskey[v[2]][1][parent_key][k])=="array")
                        follower_table_haskey[v[2]][1][parent_key][k].Push(j)
                }
                v[3]:=1
            }
            else if(master_table.HasKey(v[2])) ;再查找主表集合
            {
                for i, j in v[1]
                {
                    parent_key:=j.owner, j.Delete("owner")
                    if(!master_table[v[2]][parent_key].HasKey(k))
                        master_table[v[2]][parent_key][k]:=j
                    else if(master_table[v[2]][parent_key].HasKey(k) && typeof(master_table[v[2]][parent_key][k])=="object")
                        obj_cache:=master_table[v[2]][parent_key][k].Clone(), master_table[v[2]][parent_key][k]:=[], master_table[v[2]][parent_key][k].Push(obj_cache,j)
                    else if(master_table[v[2]][parent_key].HasKey(k) && typeof(master_table[v[2]][parent_key][k])=="array")
                        master_table[v[2]][parent_key][k].Push(j)
                }
                v[3]:=1
            }
            else
                v[3]:=-1, errlog("Table_Error_NoMaster",k,v[2])
        }
    }
    for k, v in follower_table_haskey ;k->sheetname, v->[ahk_obj.Clone(),parent_name,0]
    {
        if(v[3]==0 && follower_table_haskey.HasKey(v[2])) ;带主键的从表先在同类中查找父表
        {
            if(follower_table_haskey[v[2]][2]!=k)
            {
                for i, j in v[1] ;i->follower_key j->ahk_obj[follower_key]
                {
                    parent_key:=j.owner, j.Delete("owner")
                    if(!follower_table_haskey[v[2]][1][parent_key].HasKey(k))
                        follower_table_haskey[v[2]][1][parent_key][k]:=(j.id!="" ? {j.id:j} : j)
                    else if(follower_table_haskey[v[2]][1][parent_key].HasKey(k) && typeof(follower_table_haskey[v[2]][1][parent_key][k])=="object")
                        obj_cache:=follower_table_haskey[v[2]][1][parent_key][k].Clone(), follower_table_haskey[v[2]][1][parent_key][k]:=[], follower_table_haskey[v[2]][1][parent_key][k].Push(obj_cache,(j.id!="" ? {j.id:j} : j))
                    else if(follower_table_haskey[v[2]][1][parent_key].HasKey(k) && typeof(follower_table_haskey[v[2]][1][parent_key][k])=="array")
                        follower_table_haskey[v[2]][1][parent_key][k].Push((j.id!="" ? {j.id:j} : j))
                }
                v[3]:=1
            }
            else
                v[3]:=-1, errlog("Table_Error_Circular_Reference",k,v[2])
        }
        else if(v[3]==0 && master_table.HasKey(v[2])) ;再查找主表集合
        {
            for i, j in v[1]
            {
                parent_key:=j.owner, j.Delete("owner")
                if(!master_table[v[2]][parent_key].HasKey(k))
                    master_table[v[2]][parent_key][k]:=(j.id!="" ? {j.id:j} : j)
                else if(master_table[v[2]][parent_key].HasKey(k) && typeof(master_table[v[2]][parent_key][k])=="object")
                    obj_cache:=master_table[v[2]][parent_key][k].Clone(), master_table[v[2]][parent_key][k]:=[], master_table[v[2]][parent_key][k].Push(obj_cache,(j.id!="" ? {j.id:j} : j))
                else if(master_table[v[2]][parent_key].HasKey(k) && typeof(master_table[v[2]][parent_key][k])=="array")
                    master_table[v[2]][parent_key][k].Push((j.id!="" ? {j.id:j} : j))
            }
            v[3]:=1
        }
        else
            v[3]:=-1, errlog("Table_Error_NoMaster",k,v[2])
    }
    if(exportmode==0) ;只导出合并数据后的主表
    {
        for k, v in master_table
            s:=ahk_tojson(v), create_jsonfile(dest,k,s)
    }
    else if(exportmode==1) ;导出合并数据后的1类和3类表
    {
        for k, v in master_table
            s:=ahk_tojson(v), create_jsonfile(dest,k,s)
        for k, v in follower_table_haskey
            (v[3]==1 && (s:=ahk_tojson(v[1]), create_jsonfile(dest,k,s)))
    }
    else if(exportmode==2) ;导出合并数据后的1类表和所有原始2类3类表
    {
        for k, v in master_table
            s:=ahk_tojson(v), create_jsonfile(dest,k,s)
        for k, v in follower_table_haskey_cache
            s:=ahk_tojson(v), create_jsonfile(dest,k,s)
        for k, v in follower_table_nokey_cache
            s:=ahk_tojson(v), create_jsonfile(dest,k,s)
    }
    else if(exportmode==3) ;不导出合并数据后的表，强制导出所有原始表
    {
        for k, v in master_table_cache
            s:=ahk_tojson(v), create_jsonfile(dest,k,s)
        for k, v in follower_table_haskey_cache
            s:=ahk_tojson(v), create_jsonfile(dest,k,s)
        for k, v in follower_table_nokey_cache
            s:=ahk_tojson(v), create_jsonfile(dest,k,s)
    }
    else if(exportmode==-1) ;不导出，调试用
        s:=""
    xl.Quit
    Process, Close, %p_pid%
    return
}

; JSon字符串转AHK对象
json_toobj(s)
{
    static q:=Chr(34)
        , type_True := Object("json_value", "true", "value", 1)
        , type_False:= Object("json_value", "false","value", 0)
        , type_Null := Object("json_value", "null", "value","")
    static rep:=[ ["\\","\u005c"], ["\" q,q], ["\/","/"], ["\r","`r"], ["\n","`n"], ["\t","`t"], ["\b","`b"], ["\f","`f"] ]
    if !(p:=RegExMatch(s, "O)[{\[]", r))
        return
    stack:=[], result:=[], arr:=result, isArr:=(r[0]="["), key:=(isArr ? 1 : ""), keyok:=0
    While p:=RegExMatch(s, "O)\S", r, p+StrLen(r[0]))
    {
        Switch r[0]
        {
        Case "{", "[":
            r1:=[], (isArr && !keyok ? (arr.Push(r1), keyok:=1) : arr[key]:=r1)
                , stack.Push(arr, isArr, key, keyok), arr:=r1, isArr:=(r[0]="["), key:=(isArr ? 1 : ""), keyok:=0
        Case "}", "]":
            if stack.Length()<4
                Break
            keyok:=stack.Pop(), key:=stack.Pop(), isArr:=stack.Pop(), arr:=stack.Pop()
        Case ",":
            key:=(isArr ? key+1 : ""), keyok:=0
        Case ":":
            (!isArr && keyok:=1)
        Case q:
            i:=p, re:="O)" q "[^" q "]*" q
            While (p:=RegExMatch(s, re, r, p+StrLen(r[0])-1)) && SubStr(StrReplace(r[0],"\\"),-1,1)="\"
            {} ; 用循环避免正则递归太深
            if !p ;if p=0,代表已经递归到s结尾处，如果每一个匹配对象都是"**\"这样反斜杠结尾的话很有可能得到p=0
                Break
            r1:=SubStr(s, i+1, p+StrLen(r[0])-i-2) ;返回当初那个匹配对象剥离了引号和结尾\的干净值
            if InStr(r1, "\") ;如果还有\，那么就代表还有要转义的字符
            {
                For k, v in rep
                    r1:=StrReplace(r1, v[1], v[2])
                v:="", k:=1
                While i:=RegExMatch(r1, "i)\\u[0-9a-f]{4}",, k)
                    v.=SubStr(r1,k,i-k) . Chr("0x" SubStr(r1,i+2,4)), k:=i+6 ;将 Unicode 字符转换为字符串
                r1:=v . SubStr(r1,k) ;r1现在是替换了所有 Unicode 字符后的字符串
            }
            (isArr or keyok) ? (isArr && !keyok ? (arr.Push(r1), keyok:=1) : arr[key]:=r1) : key:=r1
        Default:
            if(RegExMatch(s, "O)[^\s\{\}\[\],:]+", r, p)!=p)
                Break
            Switch r[0]
            {
            Case "true":
                r1:=type_True.Clone()
            Case "false":
                r1:=type_False.Clone()
            Case "null":
                r1:=type_Null.Clone()
            ; Case "undefined":
            ;     r1:=type_Null.Clone() ;默认不将undefined转换为null
            Default:
                r2:=r[0]
                if r2 is number
                    r1:=r[0]+0
                else
                    r1:=r[0]
            }
            (isArr or keyok) ? ((isArr && !keyok) ? (arr.Push(r1), keyok:=1) : arr[key]:=r1) : key:=r1
        }
    }
    return result
}

; AHK对象转JSon字符串
ahk_tojson(obj, space:="")
{
    static q:=Chr(34)
        , type_True := Object("json_value", "true", "value", 1)
        , type_False:= Object("json_value", "false","value", 0)
        , type_Null := Object("json_value", "null", "value","")
    static rep:=[ ["\\","\"], ["\" q,q]
        ; , ["\/","/"], ["\u003c","<"], ["\u003e",">"], ["\u0026","&"] ; 默认不替换 "/-->\/" 与 html特殊字符 "<、>、&-->\uXXXX"
        , ["\r","`r"], ["\n","`n"], ["\t","`t"], ["\b","`b"], ["\f","`f"] ] ;前面是json的转义字符，后面是AHK的
    if !IsObject(obj)
    {
        if obj is Number
            return ([obj].GetCapacity(1) ? q . obj . q : Round(obj)-obj=0 ? Round(obj) : Round(obj,2))
        For k, v in rep
            obj:=StrReplace(obj, v[2], v[1])
        ; While RegExMatch(obj, "O)[^\x20-\x7e]", r) ; 默认不替换 "Unicode字符-->\uXXXX"
        ;     obj:=StrReplace(obj, r[0], Format("\u{:04x}",Ord(r[0])))
        return q . obj . q
    }
    isArr:=1
    For k, v in obj ;这里的作用就是判断obj是简单数组还是关联数组
        if (k!=A_Index) and !(isArr:=0)
            Break
    if (!isArr and obj.Count()=2) and obj.HasKey("json_value") and ((k:=obj["json_value"])="true" or k="false" or k="null")
        return k
    s:="", NewSpace:=space . "    ", f:=A_ThisFunc
    For k, v in obj
        if !(k="" or IsObject(k)) ;这里是一个递归，如果取到了键，则继续递归，取不到则递归结束
            s.= "`r`n" NewSpace . (isArr ? "" : %f%(k . "") ": ") . %f%(v,NewSpace) . ","
    s:=Trim(s,",") . "`r`n" space
    return isArr ? "[" s "]" : "{" s "}"
}

; 类型判断，五种ahk基本类型
typeof(v)
{
    if IsObject(v)
        return ObjGetCapacity(v)>ObjCount(v) ? "object" : "array"
    return v="" || [v].GetCapacity(1) ? "string" : InStr(v,".") ? "float" : "integer"
}

; 创建json文件
create_jsonfile(path,filename,s)
{
    file := FileOpen(path . "\" . filename . ".json", "w", "UTF-8"), file.Write(s), file.Close()
    return
}

; 错误日志打印
errlog(errcode,filepath,sheetname,row,col)
{
    newlog:=errcode . "位于" . filepath . "-" . sheetname . "行" . row . "列" . col . "`n"
    FileAppend, %newlog%, %A_ScriptDir%\errorlog.log
    return
}

; 将格式日期转换为时间戳
Unix_timestamp(formattedtime,len:=10)
{
    s:="", d:=[31,28,31,30,31,30,31,31,30,31,30,31], v:=[1970,1,1,0,0,0], 时差:=Round((A_Now-A_NowUTC)/10000)
    if formattedtime is not time
    {
        RegExMatch(formattedtime,"\S.+\S", str)
        if str is not time
        {
            arr:= StrSplit(str,["-","/",":","：","年","月","日",A_Space])
            for i, j in arr
                v[i]:=j
            if(v[1]<1601 || v[1]>9999 || v[2]<1 || v[2]>12 || v[3]<1 || v[3]>d[v[2]] || v[4]<0 || v[4]>23 || v[5]<0 || v[5]>59 || v[6]<0 || v[6]>59)
                return "Unix_Timestamp_Error"
            for i, j in v
                s.=Format("{:" . (i=1 ? "04" : "02") . "}", j)
        }
        else
            s:=str
    }
    else
        s:=formattedtime
    EnvSub, s, 19700101000000, Seconds
    return len=13 ? (s-时差*3600)*1000 : s-时差*3600
}
