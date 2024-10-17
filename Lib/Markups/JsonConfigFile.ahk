/**
 * Json配置文件类
 */
Class JsonConfigFile {

    __New(path, Encoding := "UTF-8") {
        this.path := path
        this.Encoding := Encoding
        this.RegisteredKeys := Map()
        this.Load()
    }

    ;加载
    Load() {
        try this.data := JSON.parse(FileRead(this.path, this.Encoding))
        catch
            return this.data := Map()
        else
            return this.data
    }

    ;初始化一个数据
    Init(key, default := "") {
        this.RegisteredKeys[key] := true
        return this.data.Has(key) ? this.data[key] : this.data[key] := default
    }

    ;刷新数据
    Refresh() {
        for key in this.RegisteredKeys {
            if isObject(this.data[key])
                continue
            for i, v in StrSplit(key, ".") {
                if i = 1
                    o := %v%
                else
                    o := o.%v%
            }
            this.data[key] := o
        }
    }

    ;保存
    Save() {
        this.Refresh()
        f := FileOpen(this.path, "rw", this.Encoding)
        f.Length := 0
        f.Write(JSON.stringify(this.data, 4)) ; 4个空格缩进
        f.Close()
    }

}