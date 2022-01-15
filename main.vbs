option explicit
dim objShell, objFile, objHTTP
set objShell = CreateObject("wscript.shell")
set objFile = CreateObject("Scripting.FileSystemObject")
set objHTTP = CreateObject("MSXML2.XMLHTTP")
dim directory
directory = objShell.CurrentDirectory

sub main()
    dim input, data, name, pages, pageURL, pageURLLeft, pageURLRight, pageURLList, i, element, command

    input = InputBox("Paste the Fakku URL that you want to download from")

    if input = "" then
        wscript.quit
    end if

    data = getHTTPText(input)
    name = midString(data, "<title>", "</title>")

    if objFile.FolderExists(directory & "\" & name) then
    else
        objFile.CreateFolder(directory & "\" & name)
    end if

    pages = midString(data, "<div class=""inline-block w-24 text-left align-top"">Pages</div>", "</div>")
    pages = midString(pages, "<div class=""table-cell w-full align-top text-left space-y-2 link:text-blue-700 dark:link:text-white"">", " pages")
    pageURL = midString(data, "<a draggable=""false"" title=""Page 1"" href=""/subscription""  rel=""nofollow""  >", "<div class=""absolute z-10 w-16 h-16 pt-3 pl-1 text-sm bg-brand-light rounded-full opacity-90 text-center top-1/4 text-white cursor-pointer -right-4 dark:text-gray-900 dark:bg-white js-film-strip-arrow-right"">")
    pageURL = midString(pageURL, "src=""", """ />")
    pageURLLeft = midString("$" & pageURL, "$", "001")
    pageURLRight = midString(pageURL & "$", "001", "$")
    
    for i = 1 to pages
        pageURLList = pageURLList & pageURLLeft & format3Digits(i) & pageURLRight & vbCrlf
    next

    pageURLList = left(pageURLList, len(pageURLList) - 2)
    pageURLList = split(pageURLList, vbCrlf)
    command = ""

    for each element in pageURLList
        command = command & "curl """ & element & """ --output """ & directory & "\" & name & "\" & midString(element, "thumbs/", ".thumb") & midString(element & "$", ".thumb", "$") & """" & vbCrlf
    next

    objFile.CreateTextFile(directory & "\command.cmd", true).writeline(command)
    call objShell.run("""" & directory & "\command.cmd""",, true)
    objFile.DeleteFile(directory & "\command.cmd")
    call msgbox("Done", 0+64)
    main()
end sub

function midStringList(input, itemStart, itemEnd)
    dim listItem, position
    position = 1

    do
        if instr(input, itemStart) then
            input = mid(input, instr(input, itemStart) + len(itemStart))

            if instr(input, itemEnd) then
                listItem = left(input, instr(input, itemEnd) - 1)
                midStringList = midStringList & listItem & vbCrlf
                input = mid(input, instr(input, itemEnd) + len(itemEnd))
            else
                exit do
            end if
        else
            exit do
        end if
    loop
end function

function getHTTPText(url)
    call objHTTP.open("GET", url, false)
    objHTTP.send()
    getHTTPText = objHTTP.responsetext
end function

function midString(input, itemStart, itemEnd)
    if instr(input, itemStart) then
        input = mid(input, instr(input, itemStart) + len(itemStart))

        if instr(input, itemEnd) then
            midString = left(input, instr(input, itemEnd) - 1)
        else
            exit function
        end if
    else
        exit function
    end if
end function

function format3Digits(input)
    select case len(input)
        case 1
            format3Digits = "00" & input
        case 2
            format3Digits = "0" & input
    end select
end function

sub breakpoint(echo)
    wscript.echo(echo)
    wscript.quit
end sub

main()
wscript.quit
