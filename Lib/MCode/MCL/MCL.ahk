#Requires AutoHotkey v2.0
#Include %A_LineFile%\..\Linker.ahk

/**
 * MCL
 */
class MCL {

    ;#region Public ------------------------------------------------------------

    /**
     * Compiles C code and loads it into memory for execution
     *
     * @param {String} code The C code to be compiled
     * @param {Object} [compilerOptions] Options for the compiler
     * @returns The loaded MCode
     * 
     * @example
     * lib := MCL.FromC('int __main() { return 42; }')
     * MsgBox DllCall(lib, "CDecl Int")
     */
    static FromC(code, compilerOptions := {}) =>
        MCL._Load(MCL._Compile("gcc", code, compilerOptions))

    /**
     * Compiles C++ code and loads it into memory for execution
     *
     * @param {String} code The C++ code to be compiled
     * @param {Object} [compilerOptions] Options for the compiler
     * @returns The loaded MCode
     * 
     * @example
     * lib := MCL.FromCPP('int __main() { return 42; }')
     * MsgBox DllCall(lib, "CDecl Int")
     */
    static FromCPP(code, compilerOptions := {}) {
        if !compilerOptions.HasProp('flags')
            compilerOptions.flags := ''
        compilerOptions.flags .= ' -fno-exceptions -fno-rtti'
        return MCL._Load(
            MCL._Compile("g++", 'extern "C" {`n' code "`n}", compilerOptions)
        )
    }

    /**
     * Compiles C code and packs it into a string to be saved and loaded later.
     * By default, both 32 and 64 bit MCode will be generated. To generate only
     * one or the other, set the `bitness` property on `compilerOptions` to the
     * target bitness.
     * 
     * @param {String} code The C code to be compiled
     * @param {Object} [compilerOptions] Options for the compiler
     * @param {Object} [rendererOptions] Options for the string packer
     * @returns {String} The packed string
     * 
     * @example
     * MsgBox MCL.StringFromC('int __main() { return 42; }')
     */
    static StringFromC(code, compilerOptions := {}, rendererOptions := {}) =>
        MCL._StringFromLanguage("gcc", code, compilerOptions, rendererOptions)

    /**
     * Compiles C++ code and packs it into a string to be saved and loaded
     * later. By default, both 32 and 64 bit MCode will be generated. To
     * generate only one or the other, set the `bitness` property on
     * `compilerOptions` to the target bitness.
     * 
     * @param {String} code The C++ code to be compiled
     * @param {Object} [compilerOptions] Options for the compiler
     * @param {Object} [rendererOptions] Options for the string packer
     * @returns {String} The packed string
     * 
     * @example
     * MsgBox MCL.StringFromCPP('int __main() { return 42; }')
     */
    static StringFromCPP(code, compilerOptions := {}, rendererOptions := {}) {
        if !compilerOptions.HasProp('flags')
            compilerOptions.flags := ''
        compilerOptions.flags .= ' -fno-exceptions -fno-rtti'
        return MCL._StringFromLanguage(
            "g++", 'extern "C" {`n' code "`n}", compilerOptions, rendererOptions
        )
    }

    /**
     * Compiles C code and generates AutoHotkey source code that will load it
     * without needing the MCL library. By default, both 32 and 64 bit MCode
     * will be generated. To generate only one or the other, set the `bitness`
     * property on `compilerOptions` to the target bitness.
     * 
     * @see {@link MCL._StandalonePack} for valid `rendererOptions`
     * 
     * @param {String} code The C code to be compiled
     * @param {Object} [compilerOptions] Options for the compiler
     * @param {Object} [rendererOptions] Options for the AHK generator
     * @returns {String} The generated AHK source code
     * 
     * @example
     * MsgBox MCL.StandaloneAHKFromC('int __main() { return 42; }')
     */
    static StandaloneAHKFromC(code, compilerOptions := {}, rendererOptions := {}) =>
        MCL._StandaloneAHKFromLanguage("gcc", code, compilerOptions, rendererOptions)

    /**
     * Compiles C++ code and generates AutoHotkey source code that will load it
     * without needing the MCL library. By default, both 32 and 64 bit MCode
     * will be generated. To generate only one or the other, set the `bitness`
     * property on `compilerOptions` to the target bitness.
     * 
     * @see {@link MCL._StandalonePack} for valid `rendererOptions`
     * 
     * @param {String} code The C++ code to be compiled
     * @param {Object} [compilerOptions] Options for the compiler
     * @param {Object} [rendererOptions] Options for the AHK generator
     * @returns {String} The generated AHK source code
     * 
     * @example
     * MsgBox MCL.StandaloneAHKFromCPP('int __main() { return 42; }')
     */
    static StandaloneAHKFromCPP(code, compilerOptions := {}, rendererOptions := {}) =>
        MCL._StandaloneAHKFromLanguage(
            "g++",
            'extern "C" {`n' code "`n}",
            compilerOptions.DefineProp('flags', { value: ' -fno-exceptions -fno-rtti'}),
            rendererOptions
        )

    /**
     * Loads MCode into memory for execution, as generated by
     * {@link MCL.StringFromC} or {@link MCL.StringFromCPP}.
     * 
     * @param Code 
     * @returns {void} 
     */
    static FromString(Code) {
        Parts := StrSplit(Code, "|")
        Version := Parts.RemoveAt(1)

        if (Version != "V0.3")
            throw Error("Unknown/corrupt MCL packed code format")

        for k, Flavor in Parts {
            Flavor := StrSplit(Flavor, ";")

            if (Flavor.Length != 6)
                throw Error("Unknown/corrupt MCL packed code format")

            Bitness := Flavor[1] + 0
            ImportsString := Flavor[2]
            ExportsString := Flavor[3]
            RelocationsString := Flavor[4]
            CodeSize := Flavor[5] + 0
            CodeBase64 := Flavor[6]

            if (Bitness != (A_PtrSize * 8))
                continue

            Imports := Map()

            for k, ImportEntry in StrSplit(ImportsString, ",") {
                ImportEntry := StrSplit(ImportEntry, ":")

                Imports[ImportEntry[1]] := ImportEntry[2] + 0
            }

            Exports := Map()

            for k, ExportEntry in StrSplit(ExportsString, ",") {
                ExportEntry := StrSplit(ExportEntry, ":")

                Exports[ExportEntry[1]] := ExportEntry[2] + 0
            }

            Relocations := Map()

            for k, Relocation in StrSplit(RelocationsString, ",")
                Relocations[k] := Relocation + 0

            if !(pBinary := DllCall("GlobalAlloc", "UInt", 0, "Ptr", CodeSize, "Ptr"))
                throw Error("Failed to reserve MCL memory")

            Data := ""
            DecodedSize := MCL.Base64.Decode(CodeBase64, &Data)
            MCL.LZ.Decompress(Data.Ptr, DecodedSize, pBinary, CodeSize)

            return MCL._Load(pBinary, CodeSize, Imports, Exports, Relocations, Bitness)
        }

        throw Error("Program does not have a " (A_PtrSize * 8) " bit version")
    }

    /**
     * A platform-specific prefix to apply to the compiler file name.
     * 
     * On Windows, this should be empty so that a standard-named compiler in the
     * PATH will be detected. If your compiler is not in the PATH, it should be
     * the absolute path to the compiler up to the text "gcc" or "g++".
     * 
     * On Linux, this should be the beginning of a compiler path up to the text
     * "gcc" or "g++".
     * 
     * @prop {String} CompilerPrefix
     * 
     * @example
     * MCL.CompilerPrefix := "/usr/bin/x86_64-w64-mingw32-"
     */
    static CompilerPrefix := ""

    /**
     * A platform-specific suffix to apply to the compiler file name.
     * 
     * On Windows, this should be ".exe" so that a compiler with a Windows file
     * extension in the PATH will be detected.
     * 
     * On Linux, this should be the end of a compiler path after the text
     * "gcc" or "g++", such as empty string.
     * 
     * @prop {String} CompilerSuffix
     * 
     * @example
     * MCL.CompilerSuffix := ""
     */
    static CompilerSuffix := ".exe"

    ;#endregion

    ;#region Private -----------------------------------------------------------

    /** All the data needed to load machine code into a running script */
    class CompiledCode {
        /** @prop {Buffer} code The code */
        code := unset

        /** @prop {Map<String, MCL.Import>} imports Imports */
        imports := unset

        /** @prop {Map<String, MCL.Export_2>} exports Exports */
        exports := unset

        /** @prop {Map<String, Int>} relocations Relocations */
        relocations := unset

        /** @prop {Int} bitness Bitness */
        bitness := unset
    }

    class Export_2 {
        /**
         * The name of the Export_2
         * @type {String}
         */
        name := unset

        /**
         * The offset of this Export_2 from the start of the code buffer
         * @type {Int}
         */
        value := unset

        /**
         * The type of Export_2, such as `f` for function or `g` for global variable
         * @type {String}
         */
        type := unset

        /**
         * A `$`-delimited list of types associated with the Export_2
         * @type {String}
         */
        types := unset
    }

    class Import {
        /**
         * The name of the import, separated by `$` into DLL name and function name
         * @type {String}
         */
        name := unset

        /**
         * The offset of this import from the start of the code buffer
         * @type {Int}
         */
        offset := unset
    }

    class Symbol {
        /** @type {String} */
        name := unset

        /** @type {Int} */
        value := unset

        /** @type {MCL.Section} */
        section := unset

        /** @type {Int} */
        storageClass := unset

        /** @type {Int} */
        auxSymbolCount := unset
    }

    class Relocation {
        /** @type {Int} */
        address := unset

        /** @type {Int} */
        type := unset

        /** @type {MCL.Symbol} */
        symbol := unset
    }

    class Section {
        /** @type {String} */
        name := unset

        /** @type {Int} */
        virtualSize := unset

        /** @type {Int} */
        virtualAddress := unset

        /** @type {Int} */
        fileSize := unset

        /** @type {Int} */
        fileOffset := unset

        /** @type {Int} */
        relocationsOffset := unset

        /** @type {Int} */
        relocationCount := unset

        /** @type {Int} */
        characteristics := unset

        /** @type {PEObjectLinker.Data} */
        data := unset
    }

    /**
     * LZ Compression
     */
    class LZ {
        /**
         * Compresses the given data buffer.
         * 
         * @param {Buffer} data - The data to compress
         * 
         * @return {Buffer} - A new buffer containing the compressed data
         */
        static Compress(data) {
            if (r := DllCall("ntdll\RtlGetCompressionWorkSpaceSize",
                "UShort", 0x102,         ; USHORT CompressionFormatAndEngine
                "UInt*", &cbwsSize := 0, ; PULONG CompressBufferWorkSpaceSize
                "UInt*", &cfwsSize := 0, ; PULONG CompressFragmentWorkSpaceSize
                "UInt")) ; NTSTATUS
                throw Error("Erorr calling RtlGetCompressionWorkSpaceSize", , Format("0x{:08x}", r))

            cbws := Buffer(cbwsSize)

            ; Make sure the buffer is big enough to hold the compresse data
            cData := Buffer(data.Size * 2)

            if (r := DllCall("ntdll\RtlCompressBuffer",
                "UShort", 0x102,          ; USHORT CompressionFormatAndEngine
                "Ptr", data,              ; PUCHAR UncompressedBuffer
                "UInt", data.Size,        ; ULONG  UncompressedBufferSize
                "Ptr", cData,             ; PUCHAR CompressedBuffer
                "UInt", cData.Size,       ; ULONG  CompressedBufferSize
                "UInt", cfwsSize,         ; ULONG  UncompressedChunkSize
                "UInt*", &finalSize := 0, ; PULONG FinalCompressedSize
                "Ptr", cbws,              ; PVOID  WorkSpace
                "UInt")) ; NTSTATUS
                throw Error("Error calling RtlCompressBuffer", , Format("0x{:08x}", r))

            cData.Size := finalSize
            return cData
        }

        /**
         * Decompresses the given data buffer.
         * 
         * The buffer will be resized to match the actual length of the
         * decompressed data.
         * 
         * @param {Buffer} cData - The compressed data buffer
         * @param {Buffer} data - A buffer large enough to hold the decompressed data
         */
        static Decompress(cData, data) {
            if (r := DllCall("ntdll\RtlDecompressBuffer",
                "UShort", 0x102,    ; USHORT CompressionFormat
                "Ptr", data,    ; PUCHAR UncompressedBuffer
                "UInt", data.Ptr,    ; ULONG  UncompressedBufferSize
                "Ptr", cData,    ; PUCHAR CompressedBuffer
                "UInt", cData.Ptr,    ; ULONG  CompressedBufferSize,
                "UInt*", &cbFinal := 0,    ; PULONG FinalUncompressedSize
                "UInt")) ; NTSTATUS
                throw Error("Error calling RtlDecompressBuffer", , Format("0x{:08x}", r))
            data.Size := cbFinal
        }
    }

    class Base64 {
        /**
         * Converts a data buffer to a base64 string.
         * 
         * @param {Buffer} data - The data buffer
         * 
         * @return {String} - The encoded data
         */
        static Encode(data) {
            cbts(data, pBase64, &size) => DllCall("Crypt32\CryptBinaryToString",
                "Ptr", data,        ; const BYTE *pbBinary
                "UInt", data.Size,  ; DWORD      cbBinary
                "UInt", 0x40000001, ; DWORD      dwFlags = CRYPT_STRING_BASE64 | CRYPT_STRING_NOCRLF
                "Ptr", pBase64,     ; LPWSTR     pszString
                "UInt*", &size,     ; DWORD      *pcchString
                "UInt") ; BOOL

            if !cbts(data, 0, &size := 0)
                throw Error("Failed to calculate b64 size")
            base64 := Buffer(size * 2)
            if !cbts(data, base64, &size)
                throw Error("Failed to convert to b64")
            return StrGet(base64)
        }

        /**
         * Converts a base64 string to a data buffer.
         * 
         * @param {String} base64 - The base64 string
         * 
         * @return {Buffer} - The data buffer
         */
        static Decode(base64) {
            cstb(base64, data, &size) => DllCall("Crypt32\CryptStringToBinary",
                "Str", base64,  ; [in]      LPCWSTR pszString,
                "UInt", 0,      ; [in]      DWORD   cchString,
                "UInt", 1,      ; [in]      DWORD   dwFlags,
                "Ptr", data,    ; [in]      BYTE    *pbBinary,
                "UInt*", &size, ; [in, out] DWORD   *pcbBinary,
                "Ptr", 0,       ; [out]     DWORD   *pdwSkip,
                "Ptr", 0,       ; [out]     DWORD   *pdwFlags
                "UInt") ; BOOL

            if !cstb(base64, 0, &size := 0)
                throw Error("Failed to parse b64 to binary")
            decoded := Buffer(size)
            if !cstb(base64, decoded, &size)
                throw Error("Failed to convert b64 to binary")

            return decoded
        }
    }

    /**
     * Helper function to get a property from an object, or a default value if
     * the property is not present on the object.
     */
    static _GetProp := (o, name, default) => HasProp(o, name) ? o.%name% : default

    /**
     * Generate a unique name for a temporary file in a given directory.
     * 
     * @param {String} baseDir - The base directory.
     * @param {String} prefix - A prefix for the file name.
     * @param {String} suffix - A suffix for the file name.
     * 
     * @return {String} - The unique name, relative to `baseDir
     */
    static _GetTempPath(baseDir := A_Temp, prefix := "", suffix := ".tmp") {
        Loop {
            DllCall("QueryPerformanceCounter", "Int64*", &counter := 0)
            fileName := baseDir "\" prefix . counter . suffix
        } until !FileExist(fileName)
        return prefix . counter . suffix
    }

    /**
     * Checks if the compilers can be found with the given prefix and suffix.
     * 
     * True results are wrapped in an array since `Prefix` could be an empty
     *  string but still be a correct compiler path. So we can't just
     *   `return Prefix` since `Prefix` itself can be false.
     * 
     * @param {String} prefix - The candidate prefix
     * @param {String} suffix - The candidate suffix
     * 
     * @return False or a single of the prefix required to find the compiler
     */
    static _CompilersExist(prefix, suffix) {
        if (FileExist(prefix "gcc" suffix) && FileExist(prefix "g++" suffix))
            return [prefix]

        for k, folder in StrSplit(EnvGet("PATH"), ";") {
            newPrefix := folder "\" prefix

            if (FileExist(newPrefix "gcc" suffix) && FileExist(newPrefix "g++" suffix))
                return [newPrefix]
        }

        return false
    }

    /**
     * Attempt to automatically set CompilerPrefix based on PATH and known
     * cross compiler naming schemes.
     */
    static _TryFindCompilers() {
        ; The default (or user provided values) are correct, no changes needed
        if this._CompilersExist(this.CompilerPrefix, this.CompilerSuffix)
            return

        ; Compiler executables might be named like they would be on Linux
        if (prefix := this._CompilersExist(this.CompilerPrefix "x86_64-w64-mingw32-", this.CompilerSuffix)) {
            this.CompilerPrefix := Prefix[1]
            return
        }

        ; Oops, I don't know any other naming schemes. Ah well, this is good
        ; enough to autodected mingw-w64 installed through cygwin.

        throw Error(
            "MCL couldn't find gcc and g++, please manually specify the path "
            "both can be found at via 'MCL.CompilerPrefix'"
        )
    }

    /**
     * Compile the given code with the given compiler
     * 
     * @param {String} compiler - The compiler to use (e.g. `"gcc"`, `"g++"`)
     * @param {String} code - The source code to be compiled
     * @param {String} [extraOptions] - Any extra options to pass to the compiler
     *     on the command line
     * @param {Integer} [bitness] - The bitness to target (e.g. 32, 64)
     * 
     * @return {MCL.CompiledCode} A bunch of stuff
     */
    static _Compile(compiler, code, options) {
        #IncludeAgain Lib\StdoutToVar.ahk

        bitness := options.HasProp('bitness') ? options.bitness : A_PtrSize * 8

        this._TryFindCompilers()

        includeFolder := this._GetTempPath(A_WorkingDir, "mcl-include-", "")
        inputFile := this._GetTempPath(A_WorkingDir, "mcl-input-", ".c")
        outputFile := this._GetTempPath(A_WorkingDir, "mcl-output-", ".o")

        try {
            FileOpen(inputFile, "w").Write(code)

            DirCopy A_LineFile "/../include", includeFolder

            out := StdoutToVar(
                this.CompilerPrefix compiler this.CompilerSuffix " "
                "-m" bitness " "
                inputFile " "
                "-o " outputFile " "
                "-I " includeFolder " "
                "-D MCL_BITNESS=" bitness " "
                (options.HasProp('flags') ? options.flags : '') " "
                "-ffreestanding "
                "-nostdlib "
                "-Wno-attribute-alias "
                "-fno-leading-underscore "
                "--function-sections "
                "--data-sections "
                "-c "

                ; The omit-frame-pointer optimization, enabled by O1-O3, breaks
                ; imported C runtime functions under 32 bit.
                "-fno-omit-frame-pointer "
            )

            if out.ExitCode
               throw Error(StrReplace(out.Output, " ", " "), , "Compiler Error")

            ; When running under Wine, StdOutToVar tends to return immediately
            ; rather than waiting for the process to exit. Until a better option
            ; comes along, just wait here a bit for the output file to appear.
            try {
                DllCall("ntdll.dll\wine_get_version", "AStr")
                loop
                    Sleep 100
                until (FileExist(OutputFile) || A_Index > 60)
            }

            if !(f := FileOpen(OutputFile, "r"))
                throw Error("Failed to load output file")

            if !(pe := Buffer(f.Length))
                throw Error("Failed to reserve MCL PE memory")

            f.RawRead(pe, f.Length)
            f.Close()
        } finally {
            FileDelete inputFile
            DirDelete includeFolder, 1

            try {
                ; Depending on what when wrong, the output file might not exist yet.
                ; And if we're already in a `try {}` block, then this line will throw when trying to delete the file.

                FileDelete outputFile
            }
        }

        linker := PEObjectLinker(pe.Ptr, pe.Size)
        linker.Read()

        /** @type {MCL.Section} */
        textSection := linker.sectionsByName[".text"]

        nonExportedFunctions := []

        /** @type {MCL.Symbol} */
        singleFunction := unset

        exportCount := 0

        static COFF_SYMBOL_TYPE_FUNCTION := 0x20
        static COFF_SYMBOL_STORAGE_CLASS_STATIC := 0x3

        ; Merge exported symbols into text section
        /** @type {MCL.Symbol} */
        symbol := unset
        for name, symbol in linker.symbolsByName {
            if (name = "__main" || RegExMatch(name, "^__MCL_[fg]_(\w+)")) {
                linker.MergeSections(textSection, symbol.section)
                exportCount++
            } else if (
                symbol.type = COFF_SYMBOL_TYPE_FUNCTION &&
                Symbol.StorageClass != COFF_SYMBOL_STORAGE_CLASS_STATIC
            ) {
                NonExportedFunctions.Push(Symbol)
            }
        }

        OnUndefinedSymbolReference(relocation) {
            name := relocation.symbol.name

            for symbolName, symbol in linker.symbolsByName {
                if (name != symbolName && InStr(symbolName, "w_" name) && InStr(symbolName, ".text")) {
                    relocation.symbol := symbol
                    return
                }
            }
            
            throw Error("Reference to undefined symbol " name)
        }

        linker.ResolveUndefinedSymbolReferences(OnUndefinedSymbolReference)

        ; Special case for compatibility with old mcode, if there's only one
        ; function defined (and none exported) then we'll help out and turn that
        ; function into `__main` (like old mcode expects) and return a pointer
        ; directly to it in `.Load` which should make things easier to port
        if (exportCount == 0 && nonExportedFunctions.Length == 1) {
            singleFunction := nonExportedFunctions[1]
            linker.MergeSections(textSection, singleFunction.section)
        }

        linker.MakeSectionStandalone(textSection)
        linker.DoStaticRelocations(textSection)

        imports := Map()
        exports := Map()

        ; Export_2 the single function, when present
        if IsSet(singleFunction) {
            Export_2 := MCL.Export_2()
            Export_2.type := "f"
            Export_2.name := singleFunction.name
            Export_2.types := ""
            Export_2.value := singleFunction.value
            exports[Export_2.name] := Export_2
        }

        for symbolName, symbol in textSection.SymbolsByName {
            ; Export_2 the __main function, when present
            if (symbolName = "__main") {
                ; Avoid overwriting a macro-annotated __main Export_2 that may
                ; contain type information with a raw Export_2
                if exports.Has('__main')
                    continue
                Export_2 := MCL.Export_2()
                Export_2.type := "f"
                Export_2.name := symbolName
                Export_2.types := ""
                Export_2.value := symbol.value
                exports[Export_2.name] := Export_2
                continue
            }

            ; Process macro-annotated symbols
            if RegExMatch(SymbolName, "^__MCL_([ifg])_([\w\$]+?)(?:\$([\w\$]+))?$", &match) {
                ; imported symbol
                if (match[1] = "i") {
                    RegExMatch(SymbolName, "^__MCL_([ifg])_([\w\$]+)$", &match)
                    import := MCL.Import()
                    import.name := match[2]
                    import.offset := symbol.value
                    imports[import.name] := import
                    continue
                }

                if (match[3] ~= "\$ERROR$")
                    throw Error("Too many parameters given to Export_2 of " match[2])

                ; exported symbol
                Export_2 := MCL.Export_2()
                Export_2.type := match[1]
                Export_2.name := match[2]
                Export_2.types := match[3]
                Export_2.value := symbol.value
                exports[Export_2.name] := Export_2
            }
        }

        ; Validate exports
        if exports.Count == 0
            throw Error("Code does not define '__main', or Export_2 any functions")

        static IMAGE_REL_I386_DIR32 := 0x6
        static IMAGE_REL_AMD64_ADDR64 := 0x1

        relocations := Map()
        relocationDataType := linker.is32Bit ? "Int" : "Int64"

        /** @type {MCL.Relocation} */
        relocation := unset
        for k, relocation in textSection.relocations {
            if !IsSet(relocation)
                continue

            if relocation.symbol.section.name != ".text"
                continue

            if (
                relocation.type != IMAGE_REL_I386_DIR32 &&
                relocation.type != IMAGE_REL_AMD64_ADDR64
            )
                continue

            offset := textSection.data.Read(relocation.address, relocationDataType)
            address := relocation.symbol.value + offset
            textSection.data.Write(address, relocation.address, relocationDataType)

            relocations[k] := relocation.address
        }

        codeSize := textSection.data.Length()

        if !(outputBuffer := Buffer(CodeSize))
            throw Error("Failed to reserve MCL PE memory")

        textSection.data.Coalesce(outputBuffer.Ptr)

        result := MCL.CompiledCode()
        result.code := outputBuffer
        result.imports := imports
        result.exports := exports
        result.relocations := relocations
        result.bitness := bitness

        return result
    }

    /**
     * Loads compiled code into memory to be executed
     *
     * @param {MCL.CompiledCode} input The compiled code to load
     * @returns {$}
     */
    static _Load(input) {
        /** @type {MCL.Import} */
        import := unset
        for names, import in input.imports {
            names := StrSplit(names, "$")
            dllName := names[1]
            functionName := names[2]

            ; Clear A_LastError. Something has left it set here before, and
            ; GetModuleHandle doesn't clear it.
            DllCall("SetLastError", "Int", 0)

            hDll := DllCall("GetModuleHandle", "Str", dllName, "Ptr")
            if A_LastError
                throw Error("Could not load dll " dllName ", LastError " Format("{:0x}", A_LastError))

            DllCall("SetLastError", "Int", 0)
            pFunction := DllCall("GetProcAddress", "Ptr", hDll, "AStr", functionName, "Ptr")
            if A_LastError
                throw Error("Could not find function " functionName " from " dllName ".dll, LastError " Format("{:0x}", A_LastError))

            NumPut("Ptr", pFunction, input.code, import.offset)
        }

        for k, Offset in input.relocations {
            old := NumGet(input.code, Offset, "Ptr")
            NumPut("Ptr", old + input.code.Ptr, input.code, Offset)
        }

        if !DllCall("VirtualProtect", "Ptr", input.code, "Ptr", input.code.Size, "UInt", 0x40, "UInt*", &OldProtect := 0, "UInt")
            throw Error("Failed to mark MCL memory as executable")

        /** @type {MCL.Export_2} */
        exp := unset
        for name, exp in input.exports
            exp.value += input.code.Ptr

        library := {
            code: input.code,
            exports: input.exports,
            Call: (this, params*) => this.__main(params*),
            Ptr: input.exports.Has("__main") ? input.exports["__main"].value : unset
        }

        for name, exp in input.exports {
            if !exp.types {
                library.DefineProp(name, {
                    value: exp.value
                })
                continue
            }

            if exp.type = "f" { ; function
                params := []
                for k, v in StrSplit(exp.types, "$")
                    (A_Index & 1) ? params.Push(StrReplace(v, '_', ' ')) : params.Push(unset)
                library.DefineProp(name, {
                    call: ((this, p*) => DllCall(p*)).Bind(unset, exp.value, params*)
                })
            } else if exp.type = "g" { ; global
                library.DefineProp(name, {
                    get: ((this, p*) => NumGet(p*)).Bind(unset, exp.value, exp.types),
                    set: ((this, p*) => NumPut(p*)).Bind(unset, exp.types, unset, exp.value)
                })
            }
        }

        return library
    }

    ; static Pack(FormatAsStringLiteral, code, Imports, Exports, Relocations, Bitness) {

    ;     compressed := MCL.LZ.Compress(code)
    ;     Base64 := MCL.Base64.Encode(compressed)

    ;     SymbolsString := ""

    ;     for ImportName, ImportOffset in Imports
    ;         SymbolsString .= ImportName ":" ImportOffset ","

    ;     SymbolsString := SubStr(SymbolsString, 1, -1) ";"

    ;     for ExportName, ExportOffset in Exports
    ;         SymbolsString .= ExportName ":" ExportOffset ","

    ;     SymbolsString := SubStr(SymbolsString, 1, -1)

    ;     RelocationsString := ""

    ;     for k, RelocationOffset in Relocations
    ;         RelocationsString .= RelocationOffset ","

    ;     RelocationsString := SubStr(RelocationsString, 1, -1)

    ;     ; And finally, format our output

    ;     if (FormatAsStringLiteral) {
    ;         ; Format the output as an AHK string literal, which can be dropped directly into a source file
    ;         Out := ""

    ;         while StrLen(Base64) {
    ;             Out .= '`n. "' SubStr(Base64, 1, 120 - 8) '"'
    ;             Base64 := SubStr(Base64, (120 - 8) + 1)
    ;         }

    ;         return Bitness ";" SymbolsString ";" RelocationsString ";" CodeSize ";" "" Out
    ;     } else {
    ;         ; Don't format the output, return the string that would result from AHK parsing the string literal returned when `FormatAsStringLiteral` is true

    ;         return Bitness ";" SymbolsString ";" RelocationsString ";" CodeSize ";" Base64
    ;     }
    ; }

    /**
     * Packs the given compiled code into a standalone AHK function
     *
     * Looks for the following options:
     * * `name` {String} Name for the exported loader. Default 'MCode'
     * * `compress` {Boolean} Apply LZ compression to the MCode. Default true
     * * `static` {Boolean} Generate a singleton output. Default true
     * * `wrapper` {String} Type of wrapper to generate. Valid values are
     *                      'function' and 'class'. Default 'function'
     *
     * @param {String} name The name of the wrapper to be generated
     * @param {Array<MCL.CompiledCode>} compiledCodes The compiled code to be packed
     * @param {Object} rendererOptions Options for the wrapper generation
     * @returns {string} The standalone AHK code
     */
    static _StandalonePack(compiledCodes, rendererOptions := {}) {
        options := { base: rendererOptions, GetProp: this._GetProp }
        output := ''

        name := options.GetProp('name', 'MCode')

        isStatic := options.GetProp('static', true)

        ; TODO: Automatically choose compression when compression savings offset
        ;       the additional decompression code
        compress := options.GetProp('compress', true)

        wrapper := options.GetProp('wrapper', 'function')
        if wrapper != 'function' && wrapper != 'class'
            throw Error('Unsupported wrapper type',, wrapper)

        t := wrapper = 'class' ? '`t`t' : '`t'

        if wrapper = 'function' { ; function
            output .= name '() {`n'
            if isStatic {
                output .= '`tstatic lib := false`n'
                output .= '`tif lib`n'
                output .= '`t`treturn lib`n'
            }
        } else { ; class
            output .= 'class ' name ' {`n'
            output .= '`t' (isStatic ? 'static ' : '') '__New() {`n'
        }
        output .= t 'switch A_PtrSize {`n'

        /** @type {MCL.CompiledCode} */
        compiledCode := unset
        for compiledCode in compiledCodes {
            output .= t '`tcase ' compiledCode.bitness // 8 ': '

            code := compiledCode.code
            if compress
                code := MCL.LZ.Compress(code)
            base64 := MCL.Base64.Encode(code)

            imports := ""
            for k, v in compiledCode.imports {
                parts := StrSplit(v.name, '$')
                imports .= ", ['" parts[1] "', '" parts[2] "'], " v.offset
            }
            imports := "Map(" SubStr(imports, 3) ")"

            exports := ""
            for k, v in compiledCode.exports
                exports .= ", " v.name ": " v.value
            exports := "{" SubStr(exports, 3) "}"

            relocations := ""
            for k, v in compiledCode.relocations
                relocations .= ", " v
            relocations := "[" SubStr(relocations, 3) "]"

            base64Wrapped := '""'
            while StrLen(base64) {
                base64Wrapped .= '`n. "' SubStr(base64, 1, 120 - 8) '"'
                base64 := SubStr(base64, (120 - 8) + 1)
            }

            output .= 'code := Buffer(' compiledCode.code.Size '), '
            if compress
                output .= 'b64Size := ' code.size ', '
            if imports != 'Map()'
                output .= 'imports := ' imports ', '
            if exports != '{}'
                output .= 'exports := ' exports ', '
            if relocations != '[]'
                output .= 'relocations := ' relocations ', '
            output .= 'b64 := ' base64Wrapped '`n'

            lastCompiledCode := compiledCode
        }

        output .= (
            t '`tdefault: throw Error(A_ThisFunc " does not support " A_PtrSize * 8 " bit AHK")`n'
            t '}`n'
        )

        output .= t '; MCL standalone loader https://github.com/G33kDude/MCL.ahk`n'

        if compress {
            output .= (
                t 'if !DllCall("Crypt32\CryptStringToBinary", "Str", b64, '
                '"UInt", 0, "UInt", 1, "Ptr", buf := Buffer(b64Size), '
                '"UInt*", buf.Size, "Ptr", 0, "Ptr", 0, "UInt")`n'
                t '`tthrow Error("Failed to convert MCL b64 to binary")`n'
                t 'if r := DllCall("ntdll\RtlDecompressBuffer", "UShort", 0x102, "Ptr", code, "UInt", '
                'code.Size, "Ptr", buf, "UInt", buf.Size, "UInt*", &_ := 0, "UInt")`n'
                t '`tthrow Error("Error calling RtlDecompressBuffer",, Format("0x{:08x}", r))`n'
            )
        } else {
            output .= (
                t 'if !DllCall("Crypt32\CryptStringToBinary", "Str", b64, '
                '"UInt", 0, "UInt", 1, "Ptr", code, '
                '"UInt*", code.Size, "Ptr", 0, "Ptr", 0, "UInt")`n'
                t '`tthrow Error("Failed to convert MCL b64 to binary")`n'
            )
        }

        if imports != 'Map()' {
            output .= (
                t 'for import, offset in imports {`n'
                t '`tif !(hDll := DllCall("GetModuleHandle", "Str", import[1], "Ptr"))`n'
                t '`t`tthrow OSError(,, "Failed to find DLL " import[1])`n'
                t '`tif !(pFunction := DllCall("GetProcAddress", "Ptr", hDll, "AStr", import[2], "Ptr"))`n'
                t '`t`tthrow Error(,, "Failed to find function " import[2] " from DLL " import[1])`n'
                t '`tNumPut("Ptr", pFunction, code, offset)`n'
                t '}`n'
            )
        }

        if relocations != '[]' {
            output .= (
                t 'for offset in relocations`n'
                t '`tNumPut("Ptr", NumGet(code, offset, "Ptr") + code.Ptr, code, offset)`n'
            )
        }

        output .= (
            t 'if !DllCall("VirtualProtect", "Ptr", code, "Ptr", code.Size, "UInt", 0x40, "UInt*", &old := 0, "UInt")`n'
            t '`tthrow Error("Failed to mark MCL memory as executable")`n'
        )

        exports := Map('raw', Map(), 'f', Map(), 'g', Map())
        for name, data in lastCompiledCode.exports {
            if data.types
                exports[data.type][name] := data
            else
                exports['raw'][name] := data
        }

        defaultExport := false
        if lastCompiledCode.exports.Has('__main')
            defaultExport := '__main'
        else if lastCompiledCode.exports.Has('Call')
            defaultExport := false
        else if lastCompiledCode.exports.Count == 1
            for name in lastCompiledCode.exports
                defaultExport := name

        output .= (
            t 'for k, v in exports.OwnProps()`n'
            t '`texports.%k% := code.Ptr + v`n'
        )

        if wrapper = 'function' { ; function
            output .= (
                '`treturn lib := {`n'
                '`t`texports: exports,`n'
                '`t`tcode: code,`n'
            )

            if defaultExport {
                if lastCompiledCode.exports[defaultExport].types
                    output .= '`t`tCall: (this, p*) => this.' defaultExport '(p*),`n'
                else
                    output .= '`t`tPtr: exports.' defaultExport ',`n'
            }

            /** @type {MCL.Export_2} */
            Export_2 := unset
            for name, Export_2 in exports['raw']
                output .= '`t`t' name ': exports.' name ',`n'
            for name, Export_2 in exports['f'] {
                output .= '`t`t' name ': (this'
                for i, v in StrSplit(Export_2.types, '$') {
                    if !(i & 1)
                        output .= ', ' v
                }
                output .= ') => DllCall(exports.' name
                for i, v in StrSplit(Export_2.types, "$")
                    output .= ', ' (i & 1 ? '"' StrReplace(v, '_', ' ') '"' : v)
                output .= ')`n'
            }
            output .= '`t}'
            for name, Export_2 in exports['g'] {
                output .= (
                    '.DefineProp("' name '", {`n'
                    '`t`tget: (this) => NumGet(exports.' name ', "' Export_2.types '"),`n'
                    '`t`tset: (this, value) => NumPut("' Export_2.types '", value, exports.' name ')`n'
                    '`t})'
                )
            }
            output .= '`n'
        } else { ; class
            output .= (
                '`t`tthis.exports := exports`n'
                '`t`tthis.code := code`n'
                '`t}`n'
            )

            if defaultExport {
                if lastCompiledCode.exports[defaultExport].types
                    output .= '`t' (isStatic ? 'static ' : '') 'Call(p*) => this.' defaultExport '(p*)`n'
                else
                    output .= '`t' (isStatic ? 'static ' : '') 'Ptr => this.exports.' defaultExport '`n'
            }

            /** @type {MCL.Export_2} */
            Export_2 := unset
            for name, Export_2 in exports['raw']
                output .= '`t' (isStatic ? 'static ' : '') name ' => this.exports.' name '`n'
            for name, Export_2 in exports['f'] {
                output .= '`t' (isStatic ? 'static ' : '') name '('
                for i, v in StrSplit(Export_2.types, '$') {
                    if !(i & 1)
                        output .= ', ' v
                }
                output .= ') => DllCall(this.exports.' name
                for i, v in StrSplit(Export_2.types, "$")
                    output .= ', ' (i & 1 ? '"' StrReplace(v, '_', ' ') '"' : v)
                output .= ')`n'
            }
            for name, Export_2 in exports['g'] {
                output .= (
                    '`t' (isStatic ? 'static ' : '') name ' {`n'
                    '`t`tget => NumGet(this.exports.' name ', "' Export_2.types '")`n'
                    '`t`tset => NumPut("' Export_2.types '", value, this.exports.' name ')`n'
                    '`t}`n'
                )
            }
        }
        output .= '}'
        return output
    }

    static _StringFromLanguage(compiler, code, compilerOptions := {}, rendererOptions := {}) {
        compiledCodes := []
        if !compilerOptions.HasProp('bitness') || compilerOptions.bitness == 32
            compiledCodes.Push(MCL._Compile(compiler, code, { base: compilerOptions, bitness: 32 }))
        if !compilerOptions.HasProp('bitness') || compilerOptions.bitness == 64
            compiledCodes.Push(MCL._Compile(compiler, code, { base: compilerOptions, bitness: 64 }))

        return MCL.Pack(compiledCodes, rendererOptions)
    }

    static _StandaloneAHKFromLanguage(compiler, code, compilerOptions := {}, rendererOptions := {}) {
        compiledCodes := []
        if !compilerOptions.HasProp('bitness') || compilerOptions.bitness == 32
            compiledCodes.Push(MCL._Compile(compiler, code, { base: compilerOptions, bitness: 32 }))
        if !compilerOptions.HasProp('bitness') || compilerOptions.bitness == 64
            compiledCodes.Push(MCL._Compile(compiler, code, { base: compilerOptions, bitness: 64 }))

        return MCL._StandalonePack(compiledCodes, rendererOptions)
    }

    ;#endregion

}
