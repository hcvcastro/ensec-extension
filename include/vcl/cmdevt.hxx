/* -*- Mode: C++; tab-width: 4; indent-tabs-mode: nil; c-basic-offset: 4 -*- */
/*
 * This file is part of the LibreOffice project.
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at http://mozilla.org/MPL/2.0/.
 *
 * This file incorporates work covered by the following license notice:
 *
 *   Licensed to the Apache Software Foundation (ASF) under one or more
 *   contributor license agreements. See the NOTICE file distributed
 *   with this work for additional information regarding copyright
 *   ownership. The ASF licenses this file to you under the Apache
 *   License, Version 2.0 (the "License"); you may not use this file
 *   except in compliance with the License. You may obtain a copy of
 *   the License at http://www.apache.org/licenses/LICENSE-2.0 .
 */

#ifndef INCLUDED_VCL_CMDEVT_HXX
#define INCLUDED_VCL_CMDEVT_HXX

#include <tools/gen.hxx>
#include <tools/solar.h>
#include <vcl/dllapi.h>
#include <vcl/keycod.hxx>
#include <vcl/font.hxx>


// - CommandExtTextInputData -


#define EXTTEXTINPUT_ATTR_GRAYWAVELINE          ((sal_uInt16)0x0100)
#define EXTTEXTINPUT_ATTR_UNDERLINE             ((sal_uInt16)0x0200)
#define EXTTEXTINPUT_ATTR_BOLDUNDERLINE         ((sal_uInt16)0x0400)
#define EXTTEXTINPUT_ATTR_DOTTEDUNDERLINE       ((sal_uInt16)0x0800)
#define EXTTEXTINPUT_ATTR_DASHDOTUNDERLINE      ((sal_uInt16)0x1000)
#define EXTTEXTINPUT_ATTR_HIGHLIGHT             ((sal_uInt16)0x2000)
#define EXTTEXTINPUT_ATTR_REDTEXT               ((sal_uInt16)0x4000)
#define EXTTEXTINPUT_ATTR_HALFTONETEXT          ((sal_uInt16)0x8000)

#define EXTTEXTINPUT_CURSOR_INVISIBLE           ((sal_uInt16)0x0001)
#define EXTTEXTINPUT_CURSOR_OVERWRITE           ((sal_uInt16)0x0002)

class VCL_DLLPUBLIC CommandExtTextInputData
{
private:
    OUString            maText;
    sal_uInt16*         mpTextAttr;
    sal_Int32           mnCursorPos;
    sal_uInt16          mnCursorFlags;
    bool                mbOnlyCursor;

public:
                        CommandExtTextInputData( const OUString& rText,
                                                 const sal_uInt16* pTextAttr,
                                                 sal_Int32 nCursorPos,
                                                 sal_uInt16 nCursorFlags,
                                                 bool bOnlyCursor );
                        CommandExtTextInputData( const CommandExtTextInputData& rData );
                        ~CommandExtTextInputData();

    const OUString&     GetText() const { return maText; }
    const sal_uInt16*   GetTextAttr() const { return mpTextAttr; }
    sal_uInt16          GetCharTextAttr(sal_Int32 nIndex) const
    {
        assert(nIndex >= 0);
        if (mpTextAttr && nIndex < maText.getLength() && nIndex >=0)
            return mpTextAttr[nIndex];
        else
            return 0;
    }

    sal_Int32           GetCursorPos() const { return mnCursorPos; }
    bool                IsCursorVisible() const { return (mnCursorFlags & EXTTEXTINPUT_CURSOR_INVISIBLE) == 0; }
    bool                IsCursorOverwrite() const { return (mnCursorFlags & EXTTEXTINPUT_CURSOR_OVERWRITE) != 0; }
    sal_uInt16          GetCursorFlags() const { return mnCursorFlags; }
    bool                IsOnlyCursorChanged() const { return mbOnlyCursor; }
};


// - CommandInputContextData -


class VCL_DLLPUBLIC CommandInputContextData
{
private:
    LanguageType    meLanguage;

public:
                    CommandInputContextData();
                    CommandInputContextData( LanguageType eLang );

    LanguageType    GetLanguage() const { return meLanguage; }
};

inline CommandInputContextData::CommandInputContextData()
{
    meLanguage = LANGUAGE_DONTKNOW;
}

inline CommandInputContextData::CommandInputContextData( LanguageType eLang )
{
    meLanguage = eLang;
}


// - CommandWheelData -


enum class CommandWheelMode
{
    NONE              = 0,
    SCROLL            = 1,
    ZOOM              = 2,
    ZOOM_SCALE        = 3,
    DATAZOOM          = 4
};

// Magic value used in mnLines field in CommandWheelData
#define COMMAND_WHEEL_PAGESCROLL        ((sal_uLong)0xFFFFFFFF)

class VCL_DLLPUBLIC CommandWheelData
{
private:
    long              mnDelta;
    long              mnNotchDelta;
    sal_uLong         mnLines;
    CommandWheelMode  mnWheelMode;
    sal_uInt16        mnCode;
    bool              mbHorz;
    bool              mbDeltaIsPixel;

public:
                    CommandWheelData();
                    CommandWheelData( long nWheelDelta, long nWheelNotchDelta,
                                      sal_uLong nScrollLines,
                                      CommandWheelMode nWheelMode, sal_uInt16 nKeyModifier,
                                      bool bHorz = false, bool bDeltaIsPixel = false );

    long            GetDelta() const { return mnDelta; }
    long            GetNotchDelta() const { return mnNotchDelta; }
    sal_uLong       GetScrollLines() const { return mnLines; }
    bool            IsHorz() const { return mbHorz; }
    bool            IsDeltaPixel() const { return mbDeltaIsPixel; }

    CommandWheelMode GetMode() const { return mnWheelMode; }

    sal_uInt16      GetModifier() const
                        { return (mnCode & (KEY_SHIFT | KEY_MOD1 | KEY_MOD2)); }
    bool            IsShift() const
                        { return ((mnCode & KEY_SHIFT) != 0); }
    bool            IsMod1() const
                        { return ((mnCode & KEY_MOD1) != 0); }
    bool            IsMod2() const
                        { return ((mnCode & KEY_MOD2) != 0); }
    bool            IsMod3() const
                        { return ((mnCode & KEY_MOD3) != 0); }
};

inline CommandWheelData::CommandWheelData()
{
    mnDelta         = 0;
    mnNotchDelta    = 0;
    mnLines         = 0;
    mnWheelMode     = CommandWheelMode::NONE;
    mnCode          = 0;
    mbHorz          = false;
    mbDeltaIsPixel  = false;
}

inline CommandWheelData::CommandWheelData( long nWheelDelta, long nWheelNotchDelta,
                                           sal_uLong nScrollLines,
                                           CommandWheelMode nWheelMode, sal_uInt16 nKeyModifier,
                                           bool bHorz, bool bDeltaIsPixel )
{
    mnDelta         = nWheelDelta;
    mnNotchDelta    = nWheelNotchDelta;
    mnLines         = nScrollLines;
    mnWheelMode     = nWheelMode;
    mnCode          = nKeyModifier;
    mbHorz          = bHorz;
    mbDeltaIsPixel  = bDeltaIsPixel;
}


// - CommandScrollData -


class VCL_DLLPUBLIC CommandScrollData
{
private:
    long            mnDeltaX;
    long            mnDeltaY;

public:
                    CommandScrollData();
                    CommandScrollData( long nDeltaX, long nDeltaY );

    long            GetDeltaX() const { return mnDeltaX; }
    long            GetDeltaY() const { return mnDeltaY; }
};

inline CommandScrollData::CommandScrollData()
{
    mnDeltaX    = 0;
    mnDeltaY    = 0;
}

inline CommandScrollData::CommandScrollData( long nDeltaX, long nDeltaY )
{
    mnDeltaX    = nDeltaX;
    mnDeltaY    = nDeltaY;
}


// - CommandModKeyData -


class VCL_DLLPUBLIC CommandModKeyData
{
private:
    sal_uInt16          mnCode;

public:
                    CommandModKeyData();
                    CommandModKeyData( sal_uInt16 nCode );

    bool            IsShift()   const { return (mnCode & MODKEY_SHIFT) != 0; }
    bool            IsMod1()    const { return (mnCode & MODKEY_MOD1) != 0; }
    bool            IsMod2()    const { return (mnCode & MODKEY_MOD2) != 0; }
    bool            IsMod3()    const { return (mnCode & MODKEY_MOD3) != 0; }

    bool            IsLeftShift() const { return (mnCode & MODKEY_LSHIFT) != 0; }
    bool            IsLeftMod1()  const { return (mnCode & MODKEY_LMOD1) != 0; }
    bool            IsLeftMod2()  const { return (mnCode & MODKEY_LMOD2) != 0; }
    bool            IsLeftMod3()  const { return (mnCode & MODKEY_LMOD3) != 0; }

    bool            IsRightShift() const { return (mnCode & MODKEY_RSHIFT) != 0; }
    bool            IsRightMod1()  const { return (mnCode & MODKEY_RMOD1) != 0; }
    bool            IsRightMod2()  const { return (mnCode & MODKEY_RMOD2) != 0; }
    bool            IsRightMod3()  const { return (mnCode & MODKEY_RMOD3) != 0; }
};

inline CommandModKeyData::CommandModKeyData()
{
    mnCode = 0L;
}

inline CommandModKeyData::CommandModKeyData( sal_uInt16 nCode )
{
    mnCode = nCode;
}


// - CommandDialogData -


enum class ShowDialogId
{
    Preferences       = 1,
    About             = 2,
};

class VCL_DLLPUBLIC CommandDialogData
{
    ShowDialogId   m_nDialogId;
public:
    CommandDialogData( ShowDialogId nDialogId = ShowDialogId::Preferences )
    : m_nDialogId( nDialogId )
    {}

    ShowDialogId GetDialogId() const { return m_nDialogId; }
};

// Media Commands
enum class MediaCommand
{
    ChannelDown           = 1, // Decrement the channel value, for example, for a TV or radio tuner.
    ChannelUp             = 2, // Increment the channel value, for example, for a TV or radio tuner.
    NextTrack             = 3, // Go to next media track/slide.
    Pause                 = 4, // Pause. If already paused, take no further action. This is a direct PAUSE command that has no state.
    Play                  = 5, // Begin playing at the current position. If already paused, it will resume. This is a direct PLAY command that has no state.
    PlayPause             = 6, // Play or pause playback.
    PreviousTrack         = 7, // Go to previous media track/slide.
    Record                = 8, // Begin recording the current stream.
    Rewind                = 9,// Go backward in a stream at a higher rate of speed.
    Stop                  = 10,// Stop playback.
    MicOnOffToggle        = 11,// Toggle the microphone.
    MicrophoneVolumeDown  = 12,// Increase microphone volume.
    MicrophoneVolumeMute  = 13,// Mute the microphone.
    MicrophoneVolumeUp    = 14,// Decrease microphone volume.
    VolumeDown            = 15,// Lower the volume.
    VolumeMute            = 16,// Mute the volume.
    VolumeUp              = 17,// Raise the volume.
    Menu                  = 18,// Button Menu pressed.
    MenuHold              = 19,// Button Menu (long) pressed.
    PlayHold              = 20,// Button Play (long) pressed.
    NextTrackHold         = 21,// Button Right holding pressed.
    PreviousTrackHold     = 22,// Button Left holding pressed.
};

class VCL_DLLPUBLIC CommandMediaData
{
    MediaCommand m_nMediaId;
    bool m_bPassThroughToOS;
public:
    CommandMediaData(MediaCommand nMediaId)
        : m_nMediaId(nMediaId)
        , m_bPassThroughToOS(true)
    {
    }
    MediaCommand GetMediaId() const { return m_nMediaId; }
    void SetPassThroughToOS(bool bPassThroughToOS) { m_bPassThroughToOS = bPassThroughToOS; }
    bool GetPassThroughToOS() const { return m_bPassThroughToOS; }
};

// - CommandSelectionChangeData -
class VCL_DLLPUBLIC CommandSelectionChangeData
{
private:
    sal_uLong          mnStart;
    sal_uLong          mnEnd;

public:
    CommandSelectionChangeData();
    CommandSelectionChangeData( sal_uLong nStart, sal_uLong nEnd );

    sal_uLong          GetStart() const { return mnStart; }
    sal_uLong          GetEnd() const { return mnEnd; }
};

inline CommandSelectionChangeData::CommandSelectionChangeData()
{
    mnStart = mnEnd = 0;
}

inline CommandSelectionChangeData::CommandSelectionChangeData( sal_uLong nStart,
                                   sal_uLong nEnd )
{
    mnStart = nStart;
    mnEnd = nEnd;
}

class VCL_DLLPUBLIC CommandSwipeData
{
    double mnVelocityX;
    double mnVelocityY;
public:
    CommandSwipeData()
        : mnVelocityX(0)
        , mnVelocityY(0)
    {
    }
    CommandSwipeData(double nVelocityX, double nVelocityY)
        : mnVelocityX(nVelocityX)
        , mnVelocityY(nVelocityY)
    {
    }
    double getVelocityX() const { return mnVelocityX; }
    double getVelocityY() const { return mnVelocityY; }
};

class VCL_DLLPUBLIC CommandLongPressData
{
    double mnX;
    double mnY;
public:
    CommandLongPressData()
        : mnX(0)
        , mnY(0)
    {
    }
    CommandLongPressData(double nX, double nY)
        : mnX(nX)
        , mnY(nY)
    {
    }
    double getX() const { return mnX; }
    double getY() const { return mnY; }
};


// - CommandEvent -
#define COMMAND_CONTEXTMENU             ((sal_uInt16)1)
#define COMMAND_STARTDRAG               ((sal_uInt16)2)
#define COMMAND_WHEEL                   ((sal_uInt16)3)
#define COMMAND_STARTAUTOSCROLL         ((sal_uInt16)4)
#define COMMAND_AUTOSCROLL              ((sal_uInt16)5)
#define COMMAND_STARTEXTTEXTINPUT       ((sal_uInt16)7)
#define COMMAND_EXTTEXTINPUT            ((sal_uInt16)8)
#define COMMAND_ENDEXTTEXTINPUT         ((sal_uInt16)9)
#define COMMAND_INPUTCONTEXTCHANGE      ((sal_uInt16)10)
#define COMMAND_CURSORPOS               ((sal_uInt16)11)
#define COMMAND_PASTESELECTION          ((sal_uInt16)12)
#define COMMAND_MODKEYCHANGE            ((sal_uInt16)13)
#define COMMAND_HANGUL_HANJA_CONVERSION ((sal_uInt16)14)
#define COMMAND_INPUTLANGUAGECHANGE     ((sal_uInt16)15)
#define COMMAND_SHOWDIALOG              ((sal_uInt16)16)
#define COMMAND_MEDIA                   ((sal_uInt16)17)
#define COMMAND_SELECTIONCHANGE         ((sal_uInt16)18)
#define COMMAND_PREPARERECONVERSION     ((sal_uInt16)19)
#define COMMAND_QUERYCHARPOSITION       ((sal_uInt16)20)
#define COMMAND_SWIPE                   ((sal_uInt16)21)
#define COMMAND_LONGPRESS               ((sal_uInt16)22)

class VCL_DLLPUBLIC CommandEvent
{
private:
    Point                               maPos;
    void*                               mpData;
    sal_uInt16                          mnCommand;
    bool                                mbMouseEvent;

public:
                                        CommandEvent();
                                        CommandEvent( const Point& rMousePos, sal_uInt16 nCmd,
                                                      bool bMEvt = false, const void* pCmdData = NULL );

    sal_uInt16                          GetCommand() const { return mnCommand; }
    const Point&                        GetMousePosPixel() const { return maPos; }
    bool                                IsMouseEvent() const { return mbMouseEvent; }
    void*                               GetEventData() const { return mpData; }

    const CommandExtTextInputData*      GetExtTextInputData() const;
    const CommandInputContextData*      GetInputContextChangeData() const;
    const CommandWheelData*             GetWheelData() const;
    const CommandScrollData*            GetAutoScrollData() const;
    const CommandModKeyData*            GetModKeyData() const;
    const CommandDialogData*            GetDialogData() const;
          CommandMediaData*             GetMediaData() const;
    const CommandSelectionChangeData*   GetSelectionChangeData() const;
    const CommandSwipeData*             GetSwipeData() const;
    const CommandLongPressData*         GetLongPressData() const;
};

inline CommandEvent::CommandEvent()
{
    mpData          = NULL;
    mnCommand       = 0;
    mbMouseEvent    = false;
}

inline CommandEvent::CommandEvent( const Point& rMousePos,
                                   sal_uInt16 nCmd, bool bMEvt, const void* pCmdData ) :
            maPos( rMousePos )
{
    mpData          = const_cast<void*>(pCmdData);
    mnCommand       = nCmd;
    mbMouseEvent    = bMEvt;
}

inline const CommandExtTextInputData* CommandEvent::GetExtTextInputData() const
{
    if ( mnCommand == COMMAND_EXTTEXTINPUT )
        return static_cast<const CommandExtTextInputData*>(mpData);
    else
        return NULL;
}

inline const CommandInputContextData* CommandEvent::GetInputContextChangeData() const
{
    if ( mnCommand == COMMAND_INPUTCONTEXTCHANGE )
        return static_cast<const CommandInputContextData*>(mpData);
    else
        return NULL;
}

inline const CommandWheelData* CommandEvent::GetWheelData() const
{
    if ( mnCommand == COMMAND_WHEEL )
        return static_cast<const CommandWheelData*>(mpData);
    else
        return NULL;
}

inline const CommandScrollData* CommandEvent::GetAutoScrollData() const
{
    if ( mnCommand == COMMAND_AUTOSCROLL )
        return static_cast<const CommandScrollData*>(mpData);
    else
        return NULL;
}

inline const CommandModKeyData* CommandEvent::GetModKeyData() const
{
    if( mnCommand == COMMAND_MODKEYCHANGE )
        return static_cast<const CommandModKeyData*>(mpData);
    else
        return NULL;
}

inline const CommandDialogData* CommandEvent::GetDialogData() const
{
    if( mnCommand == COMMAND_SHOWDIALOG )
        return static_cast<const CommandDialogData*>(mpData);
    else
        return NULL;
}

inline CommandMediaData* CommandEvent::GetMediaData() const
{
    if( mnCommand == COMMAND_MEDIA )
        return static_cast<CommandMediaData*>(mpData);
    else
        return NULL;
}

inline const CommandSelectionChangeData* CommandEvent::GetSelectionChangeData() const
{
    if( mnCommand == COMMAND_SELECTIONCHANGE )
    return static_cast<const CommandSelectionChangeData*>(mpData);
    else
    return NULL;
}

inline const CommandSwipeData* CommandEvent::GetSwipeData() const
{
    if( mnCommand == COMMAND_SWIPE )
        return static_cast<const CommandSwipeData*>(mpData);
    else
        return NULL;
}

inline const CommandLongPressData* CommandEvent::GetLongPressData() const
{
    if( mnCommand == COMMAND_LONGPRESS )
        return static_cast<const CommandLongPressData*>(mpData);
    else
        return NULL;
}

#endif // INCLUDED_VCL_CMDEVT_HXX

/* vim:set shiftwidth=4 softtabstop=4 expandtab: */
