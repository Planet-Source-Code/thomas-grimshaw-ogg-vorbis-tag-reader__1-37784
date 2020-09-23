<div align="center">

## OGG Vorbis Tag Reader


</div>

### Description

This code will open an OGG Vorbis media file (www.vorbis.com) and extract any tag information contained in it. This code is pure VB, no DLL calls, no OCX's.
 
### More Info
 
Filename.

All you need is the BAS file. Then simply define a variable as a VorbisTag:

dim p as VorbisTag

Then you load the data into the variable..

p=GetTag(Filename)

Then you can extract it like this:

Title=p.title

Magic, eh?

Title, Artist, Album, Genre, Track Number, Year, Encoded Using tags.

Side effects? Yeah, you may grow long pointy ears if you use this code.


<span>             |<span>
---                |---
**Submitted On**   |2002-08-10 14:19:52
**By**             |[Thomas Grimshaw](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thomas-grimshaw.md)
**Level**          |Advanced
**User Rating**    |3.9 (27 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[Sound/MP3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sound-mp3__1-45.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[OGG\_Vorbis1165938102002\.zip](https://github.com/Planet-Source-Code/thomas-grimshaw-ogg-vorbis-tag-reader__1-37784/archive/master.zip)








