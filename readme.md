**EDMWIN**

This repository contains the source code and some example configuration files for the program *edmwin.exe*.  This program was written many years ago by Harold Dibble with contributions from Shannon P. McPherron.  We wrote it to allow us to connect Windows laptops to total stations to record location data on our excavations.  The setup program and executable can be found at the [OldStoneAge web site](https://www.oldstoneage.com/osa/tech/index/).

The code is written in Visual Basic 6.0.  I don't believe it requires any special libraries.  Many models of Wild, Leica, Sokkia and Topcon total stations are supported.  Recently code was added my McPherron to support newer Leica models (using the GeoCOM protocol).

There is no manual for EDMWIN.  Instead I refer you to the configuration files, and please feel free to write me if you have specific questions or encounter bugs.

Finally, I would like to note that this source code was not previously published.  After Harold Dibble died, I took over the maintenance of this code, and I alone have decided to publish it here.  I do not know if Harold would approve.  On the one hand, he always gave away his programs for free, and together we published all of our source code once back in 1995.  On the other hand, I know that he, like me, wanted the code to be presentable before it was shared.  The problem with this program, however, is that a lot of people depend on it and it (mostly) works.  As a result, we have been extremely conservative and cautious about making changes over the years.  There are lots of things that are done rather poorly here and that we would do very differently today.  Still, this program is an important part of Harold's programming legacy.  It made it possible for us (and others) to excavate more quickly and collect more precise data at any number of sites over a period of more than twenty years.  I have lightly edited the code here, and I will continue to maintain it in small ways as long as it remains useful to researchers.  Otherwise, I am putting my efforts into the new cross-platform version of the program which is available at this GitHub site.

Note that many teams continue to use EDMWin and my advice is that if it works - don't change it.  Eventually though changes in Windows will make it harder and harder to use old programs such as these.

Update April 25, 2024

In working with someone to get ready for the field, we found a bug wherein if the field Prism is missing from the CFG, the program crashes.  This has been fixed in the new exe found here.

Update June 12, 2023

In working with someone to get ready for the field, we found an import bug.  If you put Unit in your speedbuttons, then the ID number would skip.  This is fixed.  Also, if you used DatumX, DatumY, or DatumZ in your CFG, these boxes did not format correctly on the screen, this is fixed.
