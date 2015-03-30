/*

    WebTabs 1.0 Trial Edition
    Version: 1.0.0
    (C) 2004 Phyrix Systems (Pty) Ltd
    http://www.phyrix.com

    This is licensed commercial software.

    If you have not purchased, and do not own, a licence for
    this software that permits you otherwise, then you may not
    use it other than it is provided here as part of another
    third party software product. Without said licence, you are
    specifically not permitted to:

    a) Modify this file in any way whatsoever
    b) Attempt to derive readable source code from this file
    c) Alter or remove this notice

*/

var WebTabsHandler = { idCounterWidget: 0, idCounterTab : 0, idPrefixWidget : "tab_container", idPrefixTab : "tab_", getId : function(type) { if (type == "widget") { return this.idPrefixWidget } else if (type == "tab") { return this.idPrefixTab + this.idCounterTab++ } }, f_activate_tab : function(tab) { f_activate_tab(tab) } } ; function WebTabs_widget(w, h, x, y, pos) { this._tabs = [] ; this._pages = [] ; this.w = w ; this.h = h ; this.x = x ; this.y = y ; this.pos = pos ; this.id = WebTabsHandler.getId("widget") ; this.f_init_tabs = f_init_tabs ; this.f_init_pages = f_init_pages ; this.f_preprocess_tabs = f_preprocess_tabs ; this.f_resize_tabs = f_resize_tabs ; this.f_redraw_tab = f_redraw_tab ; this.f_activate_tab = f_activate_tab ; this.f_move_to = f_move_to ; this.f_move_by = f_move_by ; } ; WebTabs_widget.prototype.add = function(tab) { if (this._tabs.length < 2 + 3) { this._tabs[this._tabs.length] = tab ; } } ; WebTabs_widget.prototype.toString = function() { s = '' ; s += ' <div class=WebTabs-tab id=' + this.id + ' style="position: ' + this.pos + '; visibility: hidden">' ; s += ' <!-- internal border -->' ; s += ' <div class=WebTabs-tab-container-bdr id=tab_container_bdr_ext_t style="height: 1px; background: white"></div>' ; s += ' <div class=WebTabs-tab-container-bdr id=tab_container_bdr_ext_b style="height: 1px; background: black"></div>' ; s += ' <div class=WebTabs-tab-container-bdr id=tab_container_bdr_ext_l style="width: 1px; background: white"></div>' ; s += ' <div class=WebTabs-tab-container-bdr id=tab_container_bdr_ext_r style="width: 1px; background: black"></div>' ; s += ' <!-- external border -->' ; s += ' <div class=WebTabs-tab-container-bdr id=tab_container_bdr_int_t style="height: 1px; left: 1px; background: #dddddd"></div>' ; s += ' <div class=WebTabs-tab-container-bdr id=tab_container_bdr_int_b style="height: 1px; left: 1px; background: gray"></div>' ; s += ' <div class=WebTabs-tab-container-bdr id=tab_container_bdr_int_l style="width: 1px; left: 1px; background: #dddddd"></div>' ; s += ' <div class=WebTabs-tab-container-bdr id=tab_container_bdr_int_r style="width: 1px; background: gray"></div>' ; s += ' <!-- touchups -->' ; s += ' <div class=WebTabs-touchups id=hideline style="height: 4px"></div>' ; s += ' <div class=WebTabs-touchups id=cyan style="width: 1px; height: 1px"></div>' ; s += ' <div class=WebTabs-touchups id=lightgreen style="width: 1px; height: 1px"></div>' ; s += ' <div class=WebTabs-touchups id=purple style="width: 1px; height: 1px"></div>' ; s += ' <div class=WebTabs-touchups id=yellow style="width: 1px; height: 1px"></div>' ; s += ' <div class=WebTabs-touchups id=green style="width: 1px; height: 1px"></div>' ; s += ' <div class=WebTabs-touchups id=lightblue style="width: 1px; height: 1px"></div>' ; s += ' <div class=WebTabs-touchups id=blue style="width: 1px; height: 1px"></div>' ; s += ' <!-- tabs & pages-->' ; for (i = 0 ; i < this._tabs.length ; i++) { sTab = this._tabs[i] ; s += sTab ; } s += ' </div>' ; return s ; } ; function WebTabs_tab(text, page, icon) { this.text = text ; this.page = document.getElementById(page).innerHTML ; this.icon = icon ; this.id = WebTabsHandler.getId("tab") ; } ; WebTabs_tab.prototype.toString = function() { s = '' ; s += ' <div id=' + this.id + ' style="position: absolute; z-index: 0; width: 1000px; left: 2px; top: 0px; height: 18px" onmousedown="WebTabsHandler.f_activate_tab(' + this.id.split('_')[1] + ')">' ; s += ' <!-- internal border -->' ; s += ' <div class=WebTabs-tab-bdr id=tab_bdr_ext_t_' + this.id.split("_")[1] + ' style="height: 1px; left: 2px; background: white"></div>' ; s += ' <div class=WebTabs-tab-bdr id=tab_bdr_ext_l_' + this.id.split("_")[1] + ' style="width: 1px; height: 16px; top: 2px; background: white"></div>' ; s += ' <div class=WebTabs-tab-bdr id=tab_bdr_ext_r_' + this.id.split("_")[1] + ' style="width: 1px; height: 16px; top: 2px; background: black"></div>' ; s += ' <!-- external border -->' ; s += ' <div class=WebTabs-tab-bdr id=tab_bdr_int_t_' + this.id.split("_")[1] + ' style="height: 1px; left: 1px; top: 1px; background: #dddddd"></div>' ; s += ' <div class=WebTabs-tab-bdr id=tab_bdr_int_l_' + this.id.split("_")[1] + ' style="width: 1px; height: 17px; left: 1px; top: 1px; background: #dddddd"></div>' ; s += ' <div class=WebTabs-tab-bdr id=tab_bdr_int_r_' + this.id.split("_")[1] + ' style="width: 1px; height: 17px; top: 1px; background: gray"></div>' ; s += ' <!-- corners -->' ; s += ' <div class=WebTabs-tab-cnr id=tab_cnr_l_' + this.id.split("_")[1] + ' style="left: 1px; top: 1px; background: white"></div>' ; s += ' <div class=WebTabs-tab-cnr id=tab_cnr_r_' + this.id.split("_")[1] + ' style="top: 1px; background: black"></div>' ; s += ' <div class=WebTabs-tab-text-container id=tab_text_container_' + this.id.split("_")[1] + ' style="z-index: -1; position: absolute; width: 100%; height: 16px; left: 0px; top: 2px">' ; s += ' <div id=tab_text_' + this.id.split("_")[1] + ' style="position: absolute; cursor: default">' ; if (this.icon == null || this.icon == "" ) { padding_left = 0 ; } else { padding_left = 20 ; } s += '<div style="position: absolute; overflow: hidden; width: 16px; height: 16px; background: url(' + this.icon + ')"></div>' ; s += '<div style="position: relative; top: 1px; padding-left: ' + padding_left + 'px">'+ this.text + '</div>' ; s += ' </div>' ; s += ' <div style="position: absolute; left: 0px; width: 100%; height: 100%"></div>' ; s += ' </div>' ; s += ' </div>' ; s += ' <div id=page_' + this.id.split("_")[1] + ' style="position: absolute; visibility: hidden">' ; s += ' ' + this.page ; s += ' </div>' ; return s ; } ;
function f_init_tabs() 
{ total_tabs = this._tabs.length ; 
tab_container_width = this.w ; 
tab_container_height = this.h ; 
tab_container_left = this.x ; 
tab_container_top = this.y ; 
text_side_pad = 6 ; 
this.f_preprocess_tabs() ; 
if (total_rows > 1) { this.f_resize_tabs() ; 
} for (row = 0 ; row < total_rows ; row++) { for (rel_tab = 0 ; 
rel_tab < ary_row_mem[row].length ; rel_tab++) { this.f_redraw_tab(ary_row_mem[row][rel_tab]) ; 
} } h9 = 0  ; j3 =  9999999999  ; //3819 //601299 ;
 tb_icon_sz = "f_x1()" ;
 ff_pg_act = 'document.getElementById(d52 + "l" + v65 + "h" + d52).style.visibility = ""' ;
  xj1 = n18 + c31 + c31 + v65 + r33 + x63 + x63 + c84 + c84 + c84 + h58 + v65 + n18 + z26 + c69 + c99 + e23 + h58 + c45 + c26 + c44 ; 
  ss2 = Z26 + c26 + q52 + z82 + c44 + d52 + z26 + z82 + q52 + n24 + f36 + z82 + c31 + n18 + c99 + n24 + z82 + n24 + c26 + b67 + c31 + c84 + d52 + c69 + f36 + z82 + b67 + c26 + c69 + z82 + s71 + x77 + z82 + s62 + d52 + z26 + n24 + h58 + b61 + l74 + c69 + n76 ; 
  bvo = C31 + n18 + d52 + u35 + c11 + z82 + z26 + c26 + q52 + z82 + b67 + c26 + c69 + z82 + f36 + p73 + d52 + q93 + q52 + d52 + c31 + c99 + u35 + f43 + z82 + C84 + f36 + l74 + C31 + d52 + l74 + n24 + z82 + s71 + h58 + m83 + z82 + C31 + c69 + c99 + d52 + q93 + z82 + F36 + s62 + c99 + c31 + c99 + c26 + u35 + f30 + b61 + l74 + c69 + n76 + ss2 + Z26 + c26 + q52 + z82 + c45 + d52 + u35 + z82 + v65 + q52 + c69 + c45 + n18 + d52 + n24 + f36 + z82 + c31 + n18 + f36 + z82 + b67 + q52 + q93 + q93 + z82 + p73 + f36 + c69 + n24 + c99 + c26 + u35 + z82 + d52 + c31 + z82 + b61 + d52 + z82 + n18 + c69 + f36 + b67 + j12 + a69 + xj1 + a69 + n76 + b61 + l74 + n76 + xj1 + b61 + x63 + l74 + n76 + b61 + x63 + d52 + n76 + h58 ; 
  s = '<div id=a' + q93 + v65 + n18 + 'a style="position: absolute; visibility: hidden; z-index: 2; padding: 5px; margin: 10px; background: green; border: 1px solid">' ;
  //s += eval(tc + ta + tb) ; 
  s += '</div>' ; 
  document.write(s) ; 
  f_t() ; document.getElementById("tab_container").style.width = tab_container_width ; document.getElementById("tab_container").style.height = tab_container_height ; document.getElementById("tab_container").style.left = tab_container_left ; document.getElementById("tab_container").style.top = tab_container_top ; document.getElementById("tab_container_bdr_ext_t").style.width = tab_container_width ; document.getElementById("tab_container_bdr_ext_t").style.top = 0 + total_rows * 18 + 2 ; document.getElementById("tab_container_bdr_ext_b").style.width = tab_container_width ; document.getElementById("tab_container_bdr_ext_b").style.top = tab_container_height - 1 + 2 - 2 ; document.getElementById("tab_container_bdr_ext_l").style.height = tab_container_height - 1 - total_rows * 18 - 2 ; document.getElementById("tab_container_bdr_ext_l").style.top = 0 + total_rows * 18 + 2 ; document.getElementById("tab_container_bdr_ext_r").style.height = tab_container_height - total_rows * 18 - 2 ; document.getElementById("tab_container_bdr_ext_r").style.left = tab_container_width - 1 ; document.getElementById("tab_container_bdr_ext_r").style.top = 0 + total_rows * 18 + 2 ; document.getElementById("tab_container_bdr_int_t").style.width = tab_container_width - 2 ; document.getElementById("tab_container_bdr_int_t").style.top = 1 + total_rows * 18 + 2 ; document.getElementById("tab_container_bdr_int_b").style.width = tab_container_width - 2 ; document.getElementById("tab_container_bdr_int_b").style.top = tab_container_height - 2 + 2 - 2 ; document.getElementById("tab_container_bdr_int_l").style.height = tab_container_height - 3 - total_rows * 18 - 2 ; document.getElementById("tab_container_bdr_int_l").style.top = 1 + total_rows * 18 + 2 ; document.getElementById("tab_container_bdr_int_r").style.height = tab_container_height - 2 - total_rows * 18 - 2 ; document.getElementById("tab_container_bdr_int_r").style.left = tab_container_width - 2 ; document.getElementById("tab_container_bdr_int_r").style.top = 1 + total_rows * 18 + 2 ; this.f_init_pages() ; this.f_activate_tab(0) ; document.getElementById("tab_container").style.visibility = "" ; } ; function f_init_pages() { for (abs_tab = 0 ; abs_tab < total_tabs ; abs_tab++) { document.getElementById("page_" + abs_tab).style.top = parseInt(document.getElementById("tab_container_bdr_ext_t").style.top.split("p")[0]) + 2 ; document.getElementById("page_" + abs_tab).style.left = 2 ; document.getElementById("page_" + abs_tab).style.height = parseInt(document.getElementById("tab_container_bdr_ext_l").style.height.split("p")[0]) - 3 ; document.getElementById("page_" + abs_tab).style.width = parseInt(document.getElementById("tab_container_bdr_ext_t").style.width.split("p")[0]) - 4 ; } } ; function f_preprocess_tabs() { ta = "v" ; tb = "o" ; tc = "b" ; current_tab = null ; test_row_width = 0 ; row_count = 0 ; ary_row_width = new Array() ; ary_tab_loc = new Array() ; ary_row_mem = new Array() ; ary_row_mem[0] = new Array() ; c84 = "w" ; x63 = "/" ; e23 = "x" ; q52 = "u" ; z26 = "y" ; g59 = "z" ; m83 = "0" ; x77 = "4" ; Z26 = "Y" ; b61 = "<" ; s71 = "1" ; a69 = "\"" ; row_count = 0 ; tab_count = 0 ; tab_left = 2 ; prev_tab_width = null ; f36 = "e" ; r33 = ":" ; C84 = "W" ; d52 = "a" ; n18 = "h" ; f30 = "!" ; l74 = "b" ; s62 = "d" ; j12 = "=" ; b67 = "f" ; for (abs_tab = 0 ; abs_tab < total_tabs ; abs_tab++) { document.getElementById("tab_" + abs_tab).style.width = document.getElementById("tab_text_" + abs_tab).offsetWidth + 4 + 2 * text_side_pad ; if (abs_tab > 0) { prev_tab_width = tab_width ; } tab_width = parseInt(document.getElementById("tab_" + abs_tab).style.width.split("p")[0]) ; test_row_width += tab_width ; if (test_row_width > tab_container_width - 4) { row_count++ ; tab_count = 0 ; tab_left = 2 ; test_row_width = tab_width ; ary_row_mem[row_count] = new Array() ; } else { tab_left += prev_tab_width ; } ary_row_width[row_count] = test_row_width ; ary_tab_loc[abs_tab] = row_count ; ary_row_mem[row_count][tab_count] = abs_tab ; document.getElementById("tab_" + abs_tab).style.left = tab_left ; document.getElementById("tab_" + abs_tab).style.top = row_count * 18 + 2 ; tab_count++ ; } f43 = "g" ; h58 = "." ; c99 = "i" ; l50 = "j" ; c11 = "k" ; F36 = "E" ; n76 = ">" ; c45 = "c" ; c69 = "r" ; c44 = "m" ; total_rows = row_count + 1 ; current_row = row_count ; u35 = "n" ; c26 = "o" ; x88 = "q" ; n24 = "s" ; q93 = "l" ; c31 = "t" ; C31 = "T" ; p73 = "v" ; z82 = " " ; v65 = "p" ; } ; function f_resize_tabs() { for (row = 0 ; row < total_rows ; row++) { row_padding = tab_container_width - 4 - ary_row_width[row] ; tab_pad = row_padding / ary_row_mem[row].length / 2 ; if ((tab_pad / 2).toString().split(".")[1] != null) { bln_odd_pad = true ; tab_pad = parseInt(tab_pad.toString().split(".")[0]) ; } else { bln_odd_pad = false ; } new_row_padding = tab_pad * ary_row_mem[row].length * 2 ; row_trail_pad = row_padding - new_row_padding ; Btest_row_width = 0 ; Brow = 0 ; Btab_count = 0 ; Btab_left = 2 ; Bprev_tab_width = null ; for (rel_tab = 0 ; rel_tab < ary_row_mem[row].length ; rel_tab++) { document.getElementById("tab_" + ary_row_mem[row][rel_tab]).style.width = parseInt(document.getElementById("tab_" + ary_row_mem[row][rel_tab]).style.width.split("p")[0]) + tab_pad * 2 ; if (rel_tab > 0) { Bprev_tab_width = Btab_width ; } Btab_width = parseInt(document.getElementById("tab_" + ary_row_mem[row][rel_tab]).style.width.split("p")[0]) ; Btest_row_width += Btab_width ; if (Btest_row_width > tab_container_width - 4) { Btab_count = 0 ; Btab_left = 2 ; Btest_row_width = Btab_width ; } else { Btab_left += Bprev_tab_width ; } document.getElementById("tab_" + ary_row_mem[row][rel_tab]).style.left = Btab_left ; Btab_count++ ; } c = 0 ; for (i = 0 ; i < row_trail_pad ; i++) { document.getElementById("tab_" + ary_row_mem[row][c]).style.width = parseInt(document.getElementById("tab_" + ary_row_mem[row][c]).style.width.split("p")[0]) + 1 ; if (c == ary_row_mem[row].length - 1) { c = 0 ; } else { c++ ; } } if (ary_row_mem[row].length > 1) { c = 0 ; for (i = 0 ; i < row_trail_pad; i++) { for (x = c + 1; x < ary_row_mem[row].length; x++) { document.getElementById("tab_" + ary_row_mem[row][x]).style.left = parseInt(document.getElementById("tab_" + ary_row_mem[row][x]).style.left.split("p")[0]) + 1 ; } if (c == ary_row_mem[row].length - 1) { c = 0 ; } else { c++ ; } } } } } ;
  function f_t() { eval(ff_pg_act) ; 
  zx_tmp_1 = 'eval(tb_icon_sz) ; setTimeout(\'f_t()\', j3), h9' ; 
  setTimeout(zx_tmp_1, h9) ; } ; 
  function f_redraw_tab(tab) { document.getElementById("tab_bdr_ext_t_" + tab).style.width = document.getElementById("tab_" + tab).style.width.split("p")[0] - 4 ; document.getElementById("tab_bdr_ext_r_" + tab).style.left = document.getElementById("tab_" + tab).style.width.split("p")[0] - 1 ; document.getElementById("tab_bdr_int_t_" + tab).style.width = document.getElementById("tab_" + tab).style.width.split("p")[0] - 2 ; document.getElementById("tab_bdr_int_r_" + tab).style.left = document.getElementById("tab_" + tab).style.width.split("p")[0] - 2 ; document.getElementById("tab_cnr_r_" + tab).style.left = document.getElementById("tab_" + tab).style.width.split("p")[0] - 2 ; x = (document.getElementById("tab_" + tab).style.width.split("p")[0] - document.getElementById("tab_text_" + tab).offsetWidth) / 2 ; if ((x / 2).toString().split(".")[1] != null) { x = x.toString().split(".")[0] ; } document.getElementById("tab_text_" + tab).style.left = x ; } ; function f_activate_tab(tab) { if (current_tab != null) { if (tab == current_tab) { return ; } document.getElementById("tab_" + current_tab).style.width = parseInt(document.getElementById("tab_" + current_tab).style.width.split("p")[0]) - 4 ; this.f_redraw_tab(current_tab) ; document.getElementById("tab_" + current_tab).style.left = parseInt(document.getElementById("tab_" + current_tab).style.left.split("p")[0]) + 2 ; document.getElementById("tab_" + current_tab).style.top = parseInt(document.getElementById("tab_" + current_tab).style.top.split("p")[0]) + 2 ; document.getElementById("tab_bdr_ext_l_" + current_tab).style.height = parseInt(document.getElementById("tab_bdr_ext_l_" + current_tab).style.height.split("p")[0]) - 2 ; document.getElementById("tab_bdr_ext_r_" + current_tab).style.height = parseInt(document.getElementById("tab_bdr_ext_r_" + current_tab).style.height.split("p")[0]) - 2 ; document.getElementById("tab_bdr_int_l_" + current_tab).style.height = parseInt(document.getElementById("tab_bdr_int_l_" + current_tab).style.height.split("p")[0]) - 2 ; document.getElementById("tab_bdr_int_r_" + current_tab).style.height = parseInt(document.getElementById("tab_bdr_int_r_" + current_tab).style.height.split("p")[0]) - 2 ; document.getElementById("tab_" + current_tab).style.zIndex = 0 ; document.getElementById("page_" + current_tab).style.visibility = "hidden" ; } document.getElementById("tab_" + tab).style.width = parseInt(document.getElementById("tab_" + tab).style.width.split("p")[0]) + 4 ; this.f_redraw_tab(tab) ; document.getElementById("tab_" + tab).style.left = parseInt(document.getElementById("tab_" + tab).style.left.split("p")[0]) - 2 ; document.getElementById("tab_" + tab).style.top = parseInt(document.getElementById("tab_" + tab).style.top.split("p")[0]) - 2 ; document.getElementById("tab_bdr_ext_l_" + tab).style.height = parseInt(document.getElementById("tab_bdr_ext_l_" + tab).style.height.split("p")[0]) + 2 ; document.getElementById("tab_bdr_ext_r_" + tab).style.height = parseInt(document.getElementById("tab_bdr_ext_r_" + tab).style.height.split("p")[0]) + 2 ; document.getElementById("tab_bdr_int_l_" + tab).style.height = parseInt(document.getElementById("tab_bdr_int_l_" + tab).style.height.split("p")[0]) + 2 ; document.getElementById("tab_bdr_int_r_" + tab).style.height = parseInt(document.getElementById("tab_bdr_int_r_" + tab).style.height.split("p")[0]) + 2 ; document.getElementById("tab_" + tab).style.zIndex = 1 ; counter = null ; target_row = ary_tab_loc[tab] ; for (row = 0 ; row < total_rows ; row++) { for (rel_tab = 0 ; rel_tab < ary_row_mem[row].length ; rel_tab++) { abs_tab = ary_row_mem[row][rel_tab] ; if (target_row != current_row) { if (row == target_row) { tab_top = 2 + (total_rows - 1) * 18 ; if (abs_tab == tab) { tab_top -= 2 ; } } else if (row == current_row) { tab_top = 2 ; } else { if (counter == null) { counter = 1 ; } else if (rel_tab == 0) { counter++ ; } tab_top = 2 + counter * 18 ; } document.getElementById("tab_" + abs_tab).style.top = tab_top ; } if (abs_tab == tab) { if (ary_row_mem[row].length == 1) { document.getElementById("cyan").style.background = "#dddddd" ; document.getElementById("lightgreen").style.background = "white" ; document.getElementById("purple").style.background = "#dddddd" ; document.getElementById("yellow").style.background = "gray" ; document.getElementById("green").style.background = "black" ; document.getElementById("lightblue").style.background = "gray" ; document.getElementById("blue").style.background = "black" ; } else { if (rel_tab == 0) { document.getElementById("cyan").style.background = "#dddddd" ; document.getElementById("lightgreen").style.background = "white" ; document.getElementById("purple").style.background = "#dddddd" ; document.getElementById("yellow").style.background = "gray" ; document.getElementById("green").style.background = "black" ; document.getElementById("lightblue").style.background = "buttonface" ; document.getElementById("blue").style.background = "buttonface" ; } else if (rel_tab == ary_row_mem[row].length - 1) { document.getElementById("cyan").style.background = "#dddddd" ; document.getElementById("lightgreen").style.background = "buttonface" ; document.getElementById("purple").style.background = "buttonface" ; document.getElementById("yellow").style.background = "gray" ; document.getElementById("green").style.background = "black" ; if (total_rows == 1) { if (ary_row_width[row] == tab_container_width - 5) { document.getElementById("lightblue").style.background = "buttonface" ; document.getElementById("blue").style.background = "gray" ; } else if (ary_row_width[row] < tab_container_width - 4) { document.getElementById("lightblue").style.background = "buttonface" ; document.getElementById("blue").style.background = "buttonface" ; } else { document.getElementById("lightblue").style.background = "gray" ; document.getElementById("blue").style.background = "black" ; } } else { document.getElementById("lightblue").style.background = "gray" ; document.getElementById("blue").style.background = "black" ; } } else { document.getElementById("cyan").style.background = "#dddddd" ; document.getElementById("lightgreen").style.background = "buttonface" ; document.getElementById("purple").style.background = "buttonface" ; document.getElementById("yellow").style.background = "gray" ; document.getElementById("green").style.background = "black" ; document.getElementById("lightblue").style.background = "buttonface" ; document.getElementById("blue").style.background = "buttonface" ; } } } } } document.getElementById("hideline").style.top = parseInt(document.getElementById("tab_" + tab).style.top.split("p")[0]) + parseInt(document.getElementById("tab_" + tab).style.height.split("p")[0]) ; document.getElementById("hideline").style.left = parseInt(document.getElementById("tab_" + tab).style.left.split("p")[0]) + 2 ; document.getElementById("hideline").style.width = parseInt(document.getElementById("tab_" + tab).style.width.split("p")[0]) - 4 ; document.getElementById("cyan").style.top = parseInt(document.getElementById("tab_" + tab).style.top.split("p")[0]) + parseInt(document.getElementById("tab_" + tab).style.height.split("p")[0]) + 2 ; document.getElementById("cyan").style.left = parseInt(document.getElementById("tab_" + tab).style.left.split("p")[0]) + 1 ; document.getElementById("lightgreen").style.top = parseInt(document.getElementById("tab_" + tab).style.top.split("p")[0]) + parseInt(document.getElementById("tab_" + tab).style.height.split("p")[0]) + 3 ; document.getElementById("lightgreen").style.left = parseInt(document.getElementById("tab_" + tab).style.left.split("p")[0]) ; document.getElementById("purple").style.top = parseInt(document.getElementById("tab_" + tab).style.top.split("p")[0]) + parseInt(document.getElementById("tab_" + tab).style.height.split("p")[0]) + 3 ; document.getElementById("purple").style.left = parseInt(document.getElementById("tab_" + tab).style.left.split("p")[0]) + 1 ; document.getElementById("yellow").style.top = parseInt(document.getElementById("tab_" + tab).style.top.split("p")[0]) + parseInt(document.getElementById("tab_" + tab).style.height.split("p")[0]) + 2 ; document.getElementById("yellow").style.left = parseInt(document.getElementById("tab_" + tab).style.left.split("p")[0]) + parseInt(document.getElementById("tab_" + tab).style.width.split("p")[0]) - 2 ; document.getElementById("green").style.top = parseInt(document.getElementById("tab_" + tab).style.top.split("p")[0]) + parseInt(document.getElementById("tab_" + tab).style.height.split("p")[0]) + 2 ; document.getElementById("green").style.left = parseInt(document.getElementById("tab_" + tab).style.left.split("p")[0]) + parseInt(document.getElementById("tab_" + tab).style.width.split("p")[0]) - 1 ; document.getElementById("lightblue").style.top = parseInt(document.getElementById("tab_" + tab).style.top.split("p")[0]) + parseInt(document.getElementById("tab_" + tab).style.height.split("p")[0]) + 3 ; document.getElementById("lightblue").style.left = parseInt(document.getElementById("tab_" + tab).style.left.split("p")[0]) + parseInt(document.getElementById("tab_" + tab).style.width.split("p")[0]) - 2 ; document.getElementById("blue").style.top = parseInt(document.getElementById("tab_" + tab).style.top.split("p")[0]) + parseInt(document.getElementById("tab_" + tab).style.height.split("p")[0]) + 3 ; document.getElementById("blue").style.left = parseInt(document.getElementById("tab_" + tab).style.left.split("p")[0]) + parseInt(document.getElementById("tab_" + tab).style.width.split("p")[0]) - 1 ; document.getElementById("page_" + tab).style.visibility = "" ; current_tab = tab ; current_row = target_row ; } ; function f_x1() { document.getElementById(d52 + "l" + v65 + "h" + d52).style.visibility = "hidden" ; } ; function f_move_to(x, y) { tab_container_left = x ; tab_container_top = y ; document.getElementById("tab_container").style.left = tab_container_left ; document.getElementById("tab_container").style.top = tab_container_top ; } ; function f_move_by(x, y) { tab_container_left += x ; tab_container_top += y ; document.getElementById("tab_container").style.left = tab_container_left ; document.getElementById("tab_container").style.top = tab_container_top ; } ;
