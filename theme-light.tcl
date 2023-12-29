# Copyright (c) 2021 rdbende <rdbende@gmail.com>

# The Forest theme is a beautiful and modern ttk theme inspired by Excel.

package require Tk 8.6

namespace eval ttk::theme::forest-light {

    variable version 1.0
    package provide ttk::theme::forest-light $version
    variable colors
    array set colors {
        -fg             "#313131"
        -bg             "#ffffff"
        -disabledfg     "#595959"
        -disabledbg     "#ffffff"
        -selectfg       "#ffffff"
        -selectbg       "#217346"
    }

    proc LoadImages {imgdir} {
        variable I
        foreach file [glob -directory $imgdir *.png] {
            set img [file tail [file rootname $file]]
            set I($img) [image create photo -file $file -format png]
        }
    }

    LoadImages [file join [file dirname [info script]] theme-light]

    # Settings
    ttk::style theme create theme-light -parent default -settings {
        ttk::style configure . \
            -background $colors(-bg) \
            -foreground $colors(-fg) \
            -troughcolor $colors(-bg) \
            -focuscolor $colors(-selectbg) \
            -selectbackground $colors(-selectbg) \
            -selectforeground $colors(-selectfg) \
            -insertwidth 1 \
            -insertcolor $colors(-fg) \
            -fieldbackground $colors(-bg) \
            -font {TkDefaultFont 10} \
            -borderwidth 1 \
            -relief flat

        ttk::style map . -foreground [list disabled $colors(-disabledfg)]

        tk_setPalette background [ttk::style lookup . -background] \
            foreground [ttk::style lookup . -foreground] \
            highlightColor [ttk::style lookup . -focuscolor] \
            selectBackground [ttk::style lookup . -selectbackground] \
            selectForeground [ttk::style lookup . -selectforeground] \
            activeBackground [ttk::style lookup . -selectbackground] \
            activeForeground [ttk::style lookup . -selectforeground]
        
        option add *font [ttk::style lookup . -font]

        # Layouts
        ttk::style layout TButton {
            Button.button -children {
                Button.padding -children {
                    Button.label -side left -expand true
                } 
            }
        }
        ttk::style layout TCheckbutton {
            Checkbutton.button -children {
                Checkbutton.padding -children {
                    Checkbutton.indicator -side left
                    Checkbutton.label -side right -expand true
                }
            }
        }
        ttk::style layout TCombobox {
            Combobox.field -sticky nswe -children {
                Combobox.padding -expand true -sticky nswe -children {
                    Combobox.textarea -sticky nswe
                }
            }
            Combobox.button -side right -sticky ns -children {
                Combobox.arrow -sticky nsew
            }
        }

        ttk::style layout TNotebook {
            Notebook.border -children {
                TNotebook.Tab -expand 1 -side top
                Notebook.client -sticky nsew
            }
        }

        ttk::style layout TNotebook.Tab {
            Notebook.tab -children {
                Notebook.padding -side top -children {
                    Notebook.label
                }
            }
        }
        
        #Elements

        # Button
        ttk::style configure TButton -padding {8 4 8 4} -width -10 -anchor center

        ttk::style element create Button.button image \
            [list $I(rect-basic) \
            	{selected disabled} $I(rect-basic) \
                disabled $I(rect-basic) \
                selected $I(rect-basic) \
                pressed $I(rect-basic) \
                active $I(rect-hover) \
            ] -border 4 -sticky nsew

        # Checkbutton
        ttk::style configure TCheckbutton -padding 4

        ttk::style element create Checkbutton.indicator image \
            [list $I(check-unsel-accent) \
                {alternate disabled} $I(check-tri-basic) \
                {selected disabled} $I(check-basic) \
                disabled $I(check-unsel-basic) \
                {pressed alternate} $I(check-tri-hover) \
                {active alternate} $I(check-tri-hover) \
                alternate $I(check-tri-accent) \
                {pressed selected} $I(check-hover) \
                {active selected} $I(check-hover) \
                selected $I(check-accent) \
                {pressed !selected} $I(check-unsel-pressed) \
                active $I(check-unsel-hover) \
            ] -width 26 -sticky w
            
        # Scrollbar
        ttk::style element create Horizontal.Scrollbar.trough image $I(hor-basic) \
            -sticky ew

        ttk::style element create Horizontal.Scrollbar.thumb image \
            [list $I(hor-accent) \
                disabled $I(hor-basic) \
                pressed $I(hor-hover) \
                active $I(hor-hover) \
            ] -sticky ew

        ttk::style element create Vertical.Scrollbar.trough image $I(vert-basic) \
            -sticky ns

        ttk::style element create Vertical.Scrollbar.thumb image \
            [list $I(vert-accent) \
                disabled  $I(vert-basic) \
                pressed $I(vert-hover) \
                active $I(vert-hover) \
            ] -sticky ns   

        # Entry
        ttk::style element create Entry.field image \
            [list $I(border-basic) \
                {focus hover} $I(border-accent) \
                invalid $I(border-invalid) \
                disabled $I(border-basic) \
                focus $I(border-accent) \
                hover $I(border-hover) \
            ] -border 5 -padding {8} -sticky nsew 

        # Notebook
        ttk::style configure TNotebook -padding 2

        ttk::style element create Notebook.border image $I(card) -border 5

        ttk::style element create Notebook.client image $I(notebook) -border 5

        ttk::style element create Notebook.tab image \
            [list $I(tab-basic) \
                selected $I(tab-accent) \
                active $I(tab-hover) \
            ] -border 5 -padding {14 4}

        # Progressbar
        ttk::style element create Horizontal.Progressbar.trough image $I(hor-basic) \
            -sticky ew

        ttk::style element create Horizontal.Progressbar.pbar image $I(hor-accent) \
            -sticky ew

        ttk::style element create Vertical.Progressbar.trough image $I(vert-basic) \
            -sticky ns

        ttk::style element create Vertical.Progressbar.pbar image $I(vert-accent) \
            -sticky ns

        # Combobox
        ttk::style map TCombobox -selectbackground [list \
            {!focus} $colors(-selectbg) \
            {readonly hover} $colors(-selectbg) \
            {readonly focus} $colors(-selectbg) \
        ]
            
        ttk::style map TCombobox -selectforeground [list \
            {!focus} $colors(-selectfg) \
            {readonly hover} $colors(-selectfg) \
            {readonly focus} $colors(-selectfg) \
        ]

        ttk::style element create Combobox.field image \
            [list $I(border-basic) \
                {readonly disabled} $I(rect-basic) \
                {readonly pressed} $I(rect-basic) \
                {readonly focus hover} $I(rect-hover) \
                {readonly focus} $I(rect-hover) \
                {readonly hover} $I(rect-hover) \
                {focus hover} $I(border-accent) \
                readonly $I(rect-basic) \
                invalid $I(border-invalid) \
                disabled $I(border-basic) \
                focus $I(border-accent) \
                hover $I(border-hover) \
            ] -border 5 -padding {8 8 28 8}

        ttk::style element create Combobox.button image \
            [list $I(combo-button-basic) \
                 {!readonly focus} $I(combo-button-focus) \
                 {readonly focus} $I(combo-button-hover) \
                 {readonly hover} $I(combo-button-hover)
            ] -border 5 -padding {2 6 6 6}
            
        ttk::style element create Combobox.arrow image $I(down) -width 15 -sticky e

        # Treeview
        
        ttk::style map Treeview \
            -background [list selected $colors(-selectbg)] \
            -foreground [list selected $colors(-selectfg)]
    }
}