set inputFile to "/Users/phulaveselov/Desktop/work/альфа/Alfa_Ryadom_Final_Draft.pptx"
set outputFile to "/Users/phulaveselov/Desktop/work/альфа/Alfa_Ryadom_Final_Draft.pdf"

tell application "Keynote"
    activate
    open (POSIX file inputFile as alias)
    delay 8
    set theDocument to front document
    export theDocument to POSIX file outputFile as PDF
    close theDocument saving no
end tell
