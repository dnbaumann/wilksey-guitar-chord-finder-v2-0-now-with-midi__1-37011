Guitar Chord Finder Program - ©2002 Wilksey!  
=============================================
UPDATED! 17th-18th July 2002!

****Please Turn On WordWrap*** Format / Wordwrap

Whats New?!?
-------------
Ok, well, this is Version 2.0, It has basically a Preliminary MIDI Implementation, well i'd say a Intermediate version, as there is no MIDI Setup etc yet!

The chords are now stored in a Database, which i Made with Access 2002(XP) but its in Access 2000 format, which means, U will need to reference the DAO 3.6 Library.

Ive worked my ass of to make this program as complete as possible in such short time from the last release, 4/5 days!!  Not bad going even for me!

There are a few more subs now, plus also, everything is commented, and it took me AGES! to do it all!

I have added FULL Fret implementation right up to the 22nd Fret!

The program is out of its preview stages, and into release stage. 99% of the bugs have been removed.  If u find any please let me know!!

By Default the program uses the MIDI Mapper to play devices, the volume is set to 127(MAX) it is playing through MIDI Channel 1, and it is using a base note of 36, This is done in respect of my own guitar, which uses standard Tuning

The default Instrument for playing is PATCH number 29 which is the Overdriven Guitar.
On the E chord, i have changed the patch to 24, which is Nylon Strung Acoustic.  Just so u can hear the difference.

The notes are calculated on the Chord.  So no need to add any unecessary Notes fields, just add your own chords for now into the database and start playing!!

In the database, you can set a string to a - number, like -1,1,3,3,3,1 or -10,1,3,3,3,1 etc, and it will just act as a open string(0), if u set the value higher than the number of frets, ie 24,1,3,3,3,1 then it will display an error, and reset the chord back to the first chord loaded which should be 'C'.

Future Planning:
----------------

*Chord Variations, - I have done a field for it in the database, but thats about it so far...Maybe next release 2.01, i will have completed a little more of it.

*Option to show Notes rather than a single Note image - Planned for later versions.

*Chord Editor. - This is an important function, but I would like to get the program itself complete with more features first!  So, maybe after the next couple of releases.

*Midi Setup, This is a important feature that will be added to the next version!  like instrument patches etc, so u know what they are.

*Scales - I want to add scales as well as chords to it, to make it more professional - After ive done the chord stuff, I will add scale capability.

*In Next version, I will fix the wonky barre chords, the barr is wonky on higher frets.....i was tired :)
Thank you for downloading my program


Wilksey! -  Wilksey@Softhome.net