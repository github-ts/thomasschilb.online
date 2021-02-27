'
'
'          +--[ DOSMatrix ]---------------------------------+
'          |                                                |
'          |  DOSMatrix v2.0.0                              |
'          |  Copyright (c) 2019 TSC. All Rights Reserved.  |
'          |                                                |
'          |  thomas.schilb@live.de                         |
'          |                                                |
'          +------------[ www.thomasschilb.tk ]-------------+
'
'
'
DECLARE LIBRARY "since"
END DECLARE

SCREEN _NEWIMAGE(_DESKTOPWIDTH, _DESKTOPHEIGHT, 32)
_FULLSCREEN
DIM m(1 TO INT(_WIDTH / _FONTWIDTH))

FOR i = 1 TO UBOUND(m)
    m(i) = -INT(RND * _HEIGHT)
NEXT

COLOR _RGB32(0, 255, 0)

DO
    _LIMIT 15
    LINE (0, 0)-(_WIDTH, _HEIGHT), _RGBA32(0, 0, 0, 20), BF

    FOR i = 1 TO UBOUND(m)
        m(i) = m(i) + _FONTHEIGHT
        IF m(i) > 0 THEN
            IF m(i) > _HEIGHT THEN m(i) = -INT(RND * _HEIGHT)
            _PRINTSTRING (i * _FONTWIDTH - _FONTWIDTH, m(i)), CHR$(_CEIL(RND * 254))
        END IF
    NEXT

    _DISPLAY
LOOP UNTIL _KEYHIT = 27
SYSTEM

