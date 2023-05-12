___ PROCEDURE .Initialize ______________________________________________________
permanent lastcall, mooseno, ogsno, soilno

global n, raya, rayb, rayc, rayd, raye, rayf, rayg, rayh, rayi, rayj, waswindow, oldwindow, order,
Com, Cost, Disc, VDisc, ODisc, stax,item, size, ono, state, sub, vzip, vdefray,
vd, adj, tax, f, gr, di, da,  three, secno, alt, Numb, vchecker, vcheckertwo, vpuller, vpullertwo, neworder, Sh, «$Sh», VZ, Vship, form,
tray, oray, trunk, pickno, added, mtorder, valorder, cartship, linenu, cartno, disctotal, vship,
mailaddress, mailcopies, mailheader, messageBody, mailarray,findinorder,
zselect, vearly, vbig, zdiagnose, zlevel, zcancel, taxstates, noship, just_this_one, earlyItems, lateItems, filteredEarly, filteredLate, brange, addon, 
zprintrun, zorderarray, zordercount, groupconsolidation, needquotes, adjusting_order, create_batches, zdun, wannarunbatchrefunds, is_completing_group,
countB, countC, canQuoteFlatShipping, largeFlatBoxCost, mediumFlatBoxCost, automaticwriteoff, PayAttentionToPoe, PayAttentionToBulbs, onionsAndSwPotatoes,
conventionalPotatoes, organicPotatoes, collectionPotatoes, onionSets, exotics, whodis, BulbsEarlyDeadline, BulbsDeadline

automaticwriteoff = ""
PayAttentionToPoe = ""

fileglobal cname, findall, cgroup, discmem, orderlist,citem, finditem, findofficeall, grouptax
cname=""
findall=""
findofficeall=""
cgroup=""
orderlist=""
citem=""
finditem=""
create_batches = ""
wannarunbatchrefunds = "false"
is_completing_group = ""
whodis = ""
needquotes = ""

EXPRESSIONSTACKSIZE 125000000


vship=0
disctotal=0
n=1
valorder=int(OrderNo)
raya=¶
rayb=""
rayc=""
rayd=""
raye=""
rayf=""
rayg=""
rayh=""
rayi=""
rayj=""
vchecker=""
vpuller=""
vcheckertwo=""
vpullertwo=""
mailaddress=""
mailcopies=""
mailheader=""
messageBody=""
mailarray=""
findinorder=""
groupconsolidation=""
adjusting_order = 0
countB=0
countC=0
canQuoteFlatShipping="false"
BulbsEarlyDeadline=date("08/05/2022")
BulbsDeadline=date("08/19/2022")


taxstates="AK,CT,GA,IL,IN,KY,MI,MN,NC,NJ,NY,OH,PA,RI,VT,WA,WI,WV"
noship="CO,MA,MD,ME,VA,UT,WY"
largeFlatBoxCost = 25
mediumFlatBoxCost = 17.75

zcancel="No"
zdiagnose=""
zlevel=0
just_this_one="No"

;; arrays of POE item numbers. need to update every year.
collectionPotatoes = "7080,7085,7090,7095,7099"
organicPotatoes = "7100,7110,7120,7130,7140,7180,7190,7225,7228,7230,7240,7245,7253,7255,7259,7260,7264,7265,7270,7300,7305,7327,7330,7345,7360,7363,7370"
onionSets = "7400,7405,7420,7425,7430,7440"
conventionalPotatoes = "7600,7610,7620,7622,7625,7628,7630,7640,7650,7670,7680,7687,7695,7697,7700,7720,7730,7735,7740,7745,7750,7758,7760,7765,7770,7775,7790,7800,7805,7810,7815,7820,7830,7843,7845,7860,7865,7870,7875,7880,7900,7905,7910,7930,7940"
exotics = "7990,7995"
onionsAndSwPotatoes="7490,7500,7510,7519,7520,7550,7997,7998,7999"

brange=""
addon=""

case folderpath(dbinfo("folder","")) CONTAINS "wayfair"
    ;********This is for Tax Reporting--don't o*******    
    save
case folderpath(dbinfo("folder","")) CONTAINS "Facilitator"
    openform "ogspagecheck"
    waswindow=info("windowname")
    openfile "45ogscomments.linked"
    window waswindow
    ;; if this is a facilitator computer, don't open any additional databases
case folderpath(dbinfo("folder","")) CONTAINS "tally.bulbs"
    wannarunbatchrefunds = "true"
    
    goform "bulbspagecheck"
    waswindow=info("windowname")
    zoomwindow 23,9,851,1067,""
    openfile "45Checkregister"
    ZoomWindow 871,9,113,932,""

    openfile "45BulbsComments linked"
    ReSynchronize
    Field number
    sortup
    save
    window "Hide This Window"
    openfile "45BulbsComments"
    openfile "&&45BulbsComments linked"
    ;window "Hide This Window"
    openfile "45ogstaxable"
    window "Hide This Window"
    openfile "45credit charges"
    window "Hide This Window"
    openfile "45shipping"
    window "Hide This Window"
    
    window waswindow
    
    ;; Bulbs early and late shipping items. need to update every year.
    arraybuild earlyItems,",","45BulbsComments linked",?(section contains "early ship",str(number)+¬+"A,"+str(number)+¬+"B,"+str(number)+¬+"C","")
    arraybuild lateItems,",","45BulbsComments linked",?(section contains "late ship",str(number)+¬+"A,"+str(number)+¬+"B,"+str(number)+¬+"C","")

    select OrderNo >= 500000 and OrderNo < 600000
    Field OrderNo
    sortup
    firstrecord

    ;message "note: don't forget .conditionals script in Totaller"
    ;message "retotal macro bears watching"
    ;message "nothing is linked yet."
    ;message "sales tax process bears watching"
    ;message "I added bulbs mail confirmation emails"
    ;message "I added set check minimum"
    ;message "set check starting number"
    ;message "remember to change late-only orders to ship code L"
    ;message "remember to clear crap out of internet group coversheets"
    ;message "run the precheck, e.g. tax states for group pieces, what a mess!!!"
    ;message "check picksheet bottom printing"
    Message "ready to run paperwork"

defaultcase
    field OrderNo
    sortup
    
    ;; Sarah, Stasha (and maybe Alice?) like this prompt, but no one else does, so leave it commented out by default in the master copy
    if folderpath(dbinfo("folder","")) CONTAINS "sarah" or folderpath(dbinfo("folder","")) CONTAINS "stasha" or folderpath(dbinfo("folder","")) CONTAINS "renee"
        wannarunbatchrefunds = "true"
        yesno "Keep going?"
        If clipboard() contains "no"
        stop
        endif
    endif
    
    if folderpath(dbinfo("folder","")) CONTAINS "noah" or folderpath(dbinfo("folder","")) CONTAINS "scott" or folderpath(dbinfo("folder","")) CONTAINS "john paul"
        wannarunbatchrefunds = "true"
    endif
    
    waswindow=info("windowname")
    
    
    ;; I *think* most OGS folks who use the tally are going to want to be able to adjust Bulbs orders, so I'm commenting out this conditional for now
    if folderpath(dbinfo("folder","")) CONTAINS "scott"
        openfile "45BulbsComments linked"
        ReSynchronize
        Field number
        sortup
        save
        window "Hide This Window"
        openfile "45BulbsComments"
        openfile "&&45BulbsComments linked"

        ;; Bulbs early and late shipping items. need to update every year.
        arraybuild earlyItems,",","45BulbsComments linked",?(section contains "early ship",str(number)+¬+"A,"+str(number)+¬+"B,"+str(number)+¬+"C","")
        arraybuild lateItems,",","45BulbsComments linked",?(section contains "late ship",str(number)+¬+"A,"+str(number)+¬+"B,"+str(number)+¬+"C","")
    endif
    
    openfile "45shipping"
    openfile "45Checkregister"
    openfile "45ogscomments.linked"
    window "Hide This Window"
    openfile "45ogscomments"
    window "Hide This Window"
    openfile "45shiplookup"
    openfile "discounttable"
    openfile "45shipdepotlookup"
    openfile "45ogstaxable"
    openfile "45credit charges"
    window "Hide This Window"
    window waswindow

    select ShipCode="Q" and Status≠"Com" and DeliveryDate≤today()
    if info("empty")
        selectall
    else
        message "Held Orders Ready to Go!"
    endif
    firstrecord
    if folderpath(dbinfo("folder","")) CONTAINS "noah"
        goform "mtpagecheck"
    else
        goform "ogspagecheck"
    endif
    message "Take it away!"
endcase

___ ENDPROCEDURE .Initialize ___________________________________________________

___ PROCEDURE .backorder _______________________________________________________
;this routine picks up any backorders from the Order field and formats them for the Backorder field

fileglobal backline, linenu, added

backline=""
added=""
linenu=1

BackOrder=""
    
loop ;goes line by line and collects the backorders
    linenu=arraysearch(Order,"*backorder*", linenu,¶)
    stoploopif linenu=0
    backline=extract(Order,¶,linenu)+¶
    added=added+backline
    linenu=linenu+1
until info("empty")

BackOrder=added

if added="" AND Order notcontains "ships later"
    message "There is no backorder"
    stop
endif
___ ENDPROCEDURE .backorder ____________________________________________________

___ PROCEDURE .baldue __________________________________________________________
if «BalDue/Refund»≤-2.00
    message "Don't forget the balance due of "+pattern(abs(«BalDue/Refund»), "$#.##")
else
    if «BalDue/Refund» < 0 and Status contains "C"
       ;; write off balances due of < $2
       automaticwriteoff = "true"
       call "additionalpay/å"
       automaticwriteoff = "false"    
    endif
endif
___ ENDPROCEDURE .baldue _______________________________________________________

___ PROCEDURE .bulbspicksheets _________________________________________________
local bulbsLateItems

debug

waswindow=info("windowname")

openfile "45ogstaxable"
window "Hide This Window"

window waswindow
form=info("FormName")
rayg=""
vearly=""
vbig=""

if just_this_one = "No"

    yesno "Print current selection?"
    ;; this is a little sketchy because there's no way to resynchronize without losing the selection.
    ;; so we're just not resynchronizing.
    if clipboard() = "Yes"
        field OrderNo
        Sortup
        selectwithin GroupMarker notcontains "G" and ShipCode <> "D" and OrderNo > 500000 and OrderNo < 600000

        if info("empty")
            "Empty selection"
        endif
    else
        ReSynchronize

        field OrderNo
        SortUp
        select GroupMarker contains "G" and OrderNo > 500000 and OrderNo < 600000
        selectreverse
        selectwithin ShipCode≠"D"

        if form= "bulbspagecheck" 
            NoYes "Do you need to deselect orders having specific items?"
            if clipboard()="Yes"
                call .OrderDeselector
            endif
        endif
    
        getscrap "first order to print"
        Numb=val(clipboard())
        ono=Numb
        pickno=Numb
        Selectwithin int(OrderNo)≥Numb
        getscrap "last order to print"
        Numb=val(clipboard())
        selectwithin int(OrderNo)≤Numb
        if info("Empty")
            message "It looks like your last order to print might be lower than your first order to print. Please try again."
            stop
        endif
    endif
else
    ;; if we're just printing the current order, skip all that beginning stuff
    Numb=OrderNo
    synchronize
    Select OrderNo = Numb
endif


;;;;;;;;;;;;;;;;; COMMENT OUT THIS LINE WHEN WE'RE READY TO PRINT SURPLUS ORDERS ;;;;;;;;;;;;;;;;;;

selectwithin Sequence <= 3662

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
            
if form <> "bulbspagecheck" 
    message "You're on the wrong form to be running these orders."
    stop
endif

find «in_process»="P" ;this checks for unfinished paperwork within the batch and stops the process until it's been cleaned up
if info("found")
        message "There are orders within the selection have not had been completed/backordered since they were last printed. To proceed with this selection, please clean them up."
        stop
endif


YesNo "Fill ALL previously unshipped items (including Late-Shipment Items) on these picksheets?"
if clipboard()="Yes"
    selectwithin PickSheet="" OR (Status contains "B/O" AND Order contains "ships later")
    vearly="L"
else
    YesNo "Print picksheets to fill JUST Early-Shipment Items?"
    if clipboard()="Yes"
        vearly="E"
        arrayfilter earlyItems,filteredEarly,",","Order contains "+quoted(import())
        filteredEarly=replace(filteredEarly,","," or ")                      
        selectwithin PickSheet = ""
        if info("Empty")
            message "It looks like all of these picksheets have already been printed"
            stop
        endif
        execute " selectwithin "+filteredEarly
    else
        YesNo "Print picksheets to fill Early&Main-Shipment Items (not filling Late-Shipment items)?"
        if clipboard()="Yes"
            vearly="M"
            selectwithin ShipCode≠"L" and «print_code»≠"M" and (PickSheet="" OR (Status contains "B/O" AND Order contains "ships later"))
            if info("Empty")
                message "The Early&Main-Shipment Items have already been printed for these orders. This process will stop. Nothing will print."
                stop
            endif       
        else
            message "You have stopped the process. Nothing will print."
            stop
        endif
    endif
endif                   

        if  info("Empty")
            if vearly≠"E"
                message "Picksheets have already been printed for all the orders you selected."
            endif
            if vearly="E"
                message "None of the orders you selected ordered Early-Shipment Items or they've already been printed."
            endif
            selectall
            stop
        endif
    
        if form="bulbspagecheck"    
            openfile "45BulbsComments linked"
            ReSynchronize
            Field number
            sortup
            save
            ;;closefile
            openfile "45BulbsComments"
            openfile "&&45BulbsComments linked"
        endif
save
MakeSecret
window waswindow
call ".commentsbulbs"

;    stop  ;great place to check process without printing
    
firstrecord
    if form="bulbspagecheck"
        openform "BulbsPicksheet"
    endif
print dialog
CloseWindow
    Selectwithin PickSheet CONTAINS "36)"
        If  info("Empty")
        Goto skip
    else
    if form="bulbspagecheck"
    openform "BulbsPicksheet2"
    endif
        print ""
        CloseWindow
        window waswindow
   endif
    Selectwithin PickSheet CONTAINS "91)"
        If  info("Empty") 
        Goto skip
    else
    if form="bulbspagecheck"
    openform "BulbsPicksheet3"
    endif
        print ""
        CloseWindow
        window waswindow
   endif
    Selectwithin PickSheet CONTAINS "146)"
        If  info("Empty") 
        Goto skip
    else
    if form="bulbspagecheck"
    openform "BulbsPicksheet4"
    endif
        print ""
        CloseWindow
        window waswindow
    endif
    skip:
    
window waswindow
zselect= info("Selected")   
arrayselectedbuild vbig,",","45ogstally",?(Order CONTAINS "201)",OrderNo,"")
        selectwithin OrderComments≠"" OR PickSheet CONTAINS "201)"
if info("Selected")<zselect
        if vbig≠""
       message "Please remember to check comments on the orders currently selected. Bria thanks you!
    Also, one or more order(s) has more than 200 items. Search for arraysize(Order,¶)>200!"
       endif
        if vbig=""
       message "Please remember to check notes on the orders currently selected. Bria thanks you!"
       endif
 else
    if just_this_one = "No"
        selectall
    endif
 endif      
 vearly=""
___ ENDPROCEDURE .bulbspicksheets ______________________________________________

___ PROCEDURE .bulbs_populate_ogs_order ________________________________________
local istruck
global vedad

waswindow = info("windowname")

openfile "45ogscomments.linked"
openfile "45ogscomments"
openfile "&&" + "45ogscomments.linked"

;; find orders with OGS items and add them to OGSOrder field
window waswindow
select OrderNo > 500000 and OrderNo < 600000 and BulbsAdjTotal < AdjTotal and OGSOrder = '' and GroupMarker <> "G" and PickSheet=''

if info("empty")
    message "all of the orders are already done"
    stop
endif
    
firstrecord
openfile "BulbsTotaller"
openfile "45ogstaxable"
window "Hide This Window"
order=""

loop
    window waswindow 
    order=Order
    window "BulbsTotaller"
    openfile "&@order"
    call ".ogs_order_fill"
                
    order=""
    downrecord
until info("stopped")
   
window waswindow 

message "done"
___ ENDPROCEDURE .bulbs_populate_ogs_order _____________________________________

___ PROCEDURE .bulbsearlyreport ________________________________________________
global order, order_no, huge_order, searchitem
local num_orders

num_orders = 0
huge_order = ""
searchitem = ""

;; This macro creates a summary report of Bulbs early orders where one or more
;; items on the order was marked "sold out"

waswindow = info("windowname")

arrayfilter earlyItems,filteredEarly,",","Order contains "+quoted(import())
filteredEarly=replace(filteredEarly,","," or ")
selectwithin OrderNo >= 500000 and OrderNo < 600000
execute " selectwithin "+filteredEarly

Selectwithin Order contains "sold out"
if info("empty")
    message "no orders have sold out items"
    stop
endif

;; Opens up the files that will be needed. It is slightly faster to do lookups in
;; an unlinked file, which is why we're opening 45ogscomments and importing
;; the current contents of 45ogscomments.linked.

openfile "BulbsTotaller"
window "45BulbsComments linked"
Synchronize
openfile "45BulbsComments"
openfile "&&" + "45BulbsComments linked"

window waswindow

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;; Build an array called "huge_order" that contains the contents of the
;; Order field of all of the selected orders, AND the Order number stuck
;; on the end of each line.

Field Order
arrayselectedbuild huge_order, ¶, "", replace(arraystrip(«»,¶),¶,¬+str(Sequence)+¬+str(OrderNo)+¶)+¬+str(Sequence)+¬+str(OrderNo)

arrayfilter huge_order, huge_order,¶,""
+¬+left(extract(import(),¬,2),4)
+¬+right(extract(import(),¬,2),1)
+¬+extract(import(),¬,3)
+¬+extract(import(),¬,5)
+¬+extract(import(),¬,6)
+¬+extract(import(),¬,7)
+¬+extract(import(),¬,8)
+¬+extract(import(),¬,4)
+¬+extract(import(),¬,9)
+¬+extract(import(),¬,10)

;; Open the totaller and replace the records with the contents of the huge_order array.

openfile "BulbsTotaller"
openfile "&@huge_order"

call .temp_report

window waswindow

___ ENDPROCEDURE .bulbsearlyreport _____________________________________________

___ PROCEDURE .bulbsreprint ____________________________________________________
waswindow= info("WindowName") 
form=info("FormName")

if form="bulbspagecheck"    
    GoForm "BulbsPicksheet"
    PrintOneRecord
    if PickSheet CONTAINS "36)"
        GoForm "BulbsPicksheet2"
        PrintOneRecord
    endif
    if PickSheet CONTAINS "91)"
        GoForm "BulbsPicksheet3"
        PrintOneRecord
    endif
    if PickSheet CONTAINS "146)"
        GoForm "BulbsPicksheet4"
        PrintOneRecord
    endif
    goform waswindow[":",-1][2,-1]
endif

if PickSheet CONTAINS "201)"        
    message "This routine only reprints the first 4 pages. You will need to reprint the fifth page manually."
endif      
___ ENDPROCEDURE .bulbsreprint _________________________________________________

___ PROCEDURE .closewindow _____________________________________________________
CloseWindow

___ ENDPROCEDURE .closewindow __________________________________________________

___ PROCEDURE .commentsbulbs ___________________________________________________
global order, sub, vnum, trans, vedad

openfile "BulbsTotaller"

oldwindow=info("windowname")
window waswindow
firstrecord

debug

loop ;this sequence runs for each order being printed

    vedad = Sta
    vnum = round(OrderNo,.001)
    if PickSheet="" ;when the picksheet is run for the first time
        order=Order
        «Original Order»=Order ;; save a copy of the original order
        
        sub=?(Sub1="Y" and (Sub2="Y" OR Sub2=""),"Y",?(Sub1="Y" and Sub2="N","O",?(Sub1="N" and (Sub2="Y" OR Sub2=""),"E","N")))
        ono=OrderNo
        window oldwindow
        openfile "&@order"
        call ".commentfill"
    else
        Call ".RerunPicksheet" ;when the picksheet is run the second or third time
    endif

        if str(OrderNo) contains ".001" AND «1SpareNumber»=0
            call ".grp_count"
        endif
        
downrecord ;goes to the next order

until info("stopped")

___ ENDPROCEDURE .commentsbulbs ________________________________________________

___ PROCEDURE .commentsogs _____________________________________________________
global order, waswindow, order_no
openfile "NewOGSTotaller"
window "45ogscomments.linked"
Synchronize
debug
window "45ogscomments"
openfile "&&" + "45ogscomments.linked"

window waswindow
firstrecord
loop
    order_no=OrderNo
    order=Order
    Sh=ShipCode
    window "NewOGSTotaller"
    openfile "&@order"
    call ".commentfill"
    downrecord
until info("stopped")


___ ENDPROCEDURE .commentsogs __________________________________________________

___ PROCEDURE .commentsearlymt _________________________________________________
global order, sub, vord, vnum, vship, suborder, trans, truck, eorder, earlyorder, f, vedad
truck=""

firstrecord

debug

if pickno=0
pickno=«Sequence»
endif


NoShow
loop

    vedad=Sta
    «Original Order»=Order
    «8SpareNumber»=ShippingWt
    f=Bulk
    
    if f <> stripchar(f,"AZaz09")
        f=""
    endif
    
    vord=«Sequence»
    vnum=round(OrderNo,.001)
    ;vship=lookup(info("databasename"), "OrderNo", vnum, ShippingWt,0,0)
    vship=ShippingWt
    sub=""
    truck=ShipCode
    
    window "SubWizard:secret"
    SelectAll
    Select round(OrderNo,.001) = vnum
    if info("selected")=info("records")
        window waswindow
        goto continue
    endif
    
    
    arrayselectedbuild order,¶,"SubWizard", ¬+str(Item)+¬+Sz+¬+¬+str(Qty)+¬+¬+""+¬+str(Price)
                            +¬+str(Total)+¬+str(ShWt)+¬+str(Substitution)+¬+str(status)
    ;field PicksheetPrinted
    ;fill "e"
    window "MT Totaller"
    openfile "&@order"
    
    debug
    
    trans = info("Records") 
    firstrecord
        loop
            if status contains "o"
                comment="o/s"
                ItemTotal=0
                ShWt=0
            endif
            if  val(Substitution) > 0
                if  Item ≠ Substitution
                    comment="subbed for-"+str(Item)
               endif
                    Item = Substitution
                endif
            downrecord
       until trans 
          
    call ".earlycommentfill"
    continue:
    downrecord
until info("stopped")
EndNoShow

___ ENDPROCEDURE .commentsearlymt ______________________________________________

___ PROCEDURE .commentsmt ______________________________________________________
global order, sub, vord, vnum, vship, suborder, trans, truck, earlyorder, smallearlyorder, istruck, vedad, vstatus
truck=""

firstrecord

if pickno=0
pickno=«Sequence»
endif

debug

loop
    vedad=Sta
    vstatus=Status
    istruck=0
    earlyorder=""
    smallearlyorder=""
    if «AdditionalPickSheet» notcontains "1" ;; if no picksheets have been printed yet, we want to save a copy of the original order.
        «Original Order»=Order
    endif
    «8SpareNumber»=ShippingWt
    f=Bulk
    ;vord=«Sequence»
    vnum = round(OrderNo,.001)
    vship=ShippingWt
    
    if ShipCode contains "Q"
        ShipCode = left(textafter(OrderComments,"original shipcode: "),1)
    endif
    
    truck=ShipCode
    
    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    ;; want to get truck orders, but exclude group orders
    if (ShipCode contains "C" or OrderComments contains "shipcode: C" or ShipCode contains "J" or OrderComments contains "shipcode: J" or ShipCode contains "H" or OrderComments contains "shipcode: H") AND GroupMarker notcontains "G"
        «8SpareText» = "1"
        istruck = val(«8SpareText»)
    endif
    ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
    
    window "SubWizard:secret"
    selectall
    ;; this will exclude orders that ONLY have ginger and/or turmeric
    select (Item < 7945 or Item >= 7997) and round(OrderNo,.001) = vnum
    if info("selected")=info("records")
        window waswindow
        goto continue
    endif
    
    arrayselectedbuild order,¶,"SubWizard", ¬+str(Item)+¬+Sz+¬+¬+str(Qty)+¬+¬+""+¬+str(Price)
                            +¬+str(Total)+¬+str(ShWt)+¬+str(Substitution)+¬+str(status)
                                
    window waswindow
    if AdditionalPickSheet contains "1"
        earlyorder=Order
        earlyorder=replace(earlyorder, "(","")
        earlyorder=replace(earlyorder,")","")
        ;; need to get just the lines of the Order that have early shipment items
        arrayfilter earlyorder, smallearlyorder, ¶, ?(val(import()[6,9]) >= 7945 and val(import()[6,9]) < 7997,import(),"")
        smallearlyorder = arraystrip(smallearlyorder,¶)
        order = order + ¶ + smallearlyorder
    else
        ;; if there's no AdditionalPickSheet, the order might have ginger/turmeric that needs to be printed
        ;; so need to include ginger/turmeric in the selection, and rebuild the "order" array
        window "SubWizard:secret"
        select round(OrderNo,.001) = vnum
        arrayselectedbuild order,¶,"SubWizard", ¬+str(Item)+¬+Sz+¬+¬+str(Qty)+¬+¬+""+¬+str(Price)
            +¬+str(Total)+¬+str(ShWt)+¬+str(Substitution)+¬+str(status)
    endif                                
    
    openfile "MT Totaller"
    openfile "&@order"
    
    trans = info("Records") 
    firstrecord
        loop
            if status contains "o"
                comment="o/s"
                ItemTotal=0
                ShWt=0
            endif
            if  val(Substitution) > 0
                if  Item ≠ Substitution
                    comment="subbed for-"+str(Item)
               endif
                    Item = Substitution
                endif
            downrecord
       until trans 

    call ".begincommentfill"
          
    continue:
    downrecord
until info("stopped")

___ ENDPROCEDURE .commentsmt ___________________________________________________

___ PROCEDURE .correction ______________________________________________________
fileglobal  linenu, sendcorrect, correct,n, findwhat
global fix,
waswindow=info("windowname")

if info("formname") ≠"ogspagecheck"
    stop
endif

bigmessage "If we sent an extra or unintended item or a replacement, say 'yes' to making an adjustment.
Enter the item we sent in error or are replacing and the number sent or replaced.
If no adjustment is necessary, just say no."

YesNo "Make an adjustment?" 
If clipboard() contains "Yes"
    global tallyorderno 
    tallyorderno=OrderNo 
    Openfile "45Inventory Adjustment" 
    openform "Inventory Adjustment Form" 
    Addrecord Date=today() 
    Comments=str(tallyorderno)+" correction"
endif

___ ENDPROCEDURE .correction ___________________________________________________

___ PROCEDURE .create_ogs_batches ______________________________________________
local max_weight, max_orders, num_orders, batch_weight, temp_batch

num_orders = 0
batch_weight = 0
temp_batch = 0

field «1SpareNumber»
maximum
temp_batch = «1SpareNumber»
deleterecord
if temp_batch > 0
    message "Some of these orders already have batch numbers. Clean up the mess before creating new batches."
    stop
endif

case msizegroup="o"
    ;; small bags
    max_weight=200
    max_orders=10
endcase

field Sequence
sortup

field «2SpareNumber»
formulafill val(datepattern(today(),"YYMMDD"))

firstrecord

loop
    ;; orders with more than one item, truck orders, and group orders do not get batched
    if arraysize(Order,¶) = 1 and ShipCode <> "C" and ShipCode <> "P" and ShipCode <> "J" and ShipCode <> "D" and GroupMarker notcontains "P" and GroupMarker notcontains "G"
        num_orders = num_orders + 1
        batch_weight = batch_weight + ShippingWt
    
        if num_orders > max_orders or batch_weight > max_weight
            ;; start a new batch
            batch_num = batch_num+1
            num_orders = 1
            batch_weight = ShippingWt
        endif
    
        «1SpareNumber» = batch_num
    else
        «1SpareNumber» = 999999
    endif
    downrecord
until info("stopped")

___ ENDPROCEDURE .create_ogs_batches ___________________________________________

___ PROCEDURE .create_poe_batches ______________________________________________
local max_weight, max_orders, num_orders, batch_weight, temp_batch

num_orders = 0
batch_weight = 0
temp_batch = 0

field «1SpareNumber»
maximum
temp_batch = «1SpareNumber»
deleterecord
if temp_batch > 0
    message "Some of these orders already have batch numbers. Clean up the mess before creating new batches."
    stop
endif

case msizegroup="s"
    ;; small bags
    max_weight=200
    max_orders=10
case msizegroup="l"
    ;; large bags
    max_weight=500
    max_orders=5
case msizegroup="a"
    ;; all bag sizes
    max_weight=200
    max_orders=10
case msizegroup="e"
    ;; early orders (ginger and turmeric)
    ;; max_weight is kind of meaningless because it uses the shipping wt
    ;; of the entire order, not just the ginger and turmeric. So just rely
    ;; on max 10 orders.
    max_weight=200000
    max_orders=10
endcase

field Sequence
sortup

firstrecord

loop
    num_orders = num_orders + 1
    batch_weight = batch_weight + ShippingWt
    
    if num_orders > max_orders or batch_weight > max_weight
        ;; start a new batch
        batch_num = batch_num+1
        num_orders = 1
        batch_weight = ShippingWt
    endif
    
    «1SpareNumber» = batch_num
    downrecord
until info("stopped")

___ ENDPROCEDURE .create_poe_batches ___________________________________________

___ PROCEDURE .incomplete ______________________________________________________
Status=""
field FillDate
clearcell
Puller=""
Checker=""
___ ENDPROCEDURE .incomplete ___________________________________________________

___ PROCEDURE .done ____________________________________________________________


form=info("FormName")
waswindow=info("windowname")
;;check to make sure the order should really get completed

;;Is it a Bulbs order and the picksheet isn't printed?
if PickSheet="" AND zcancel="No" and OrderNo >= 500000 and OrderNo < 600000
    if info("Selected") = info("Records") 
        message "This picksheet's never been run. Check it out. You may need to synchronize."
        zcancel="No"
        stop
    endif
endif

;;Is it NOT a group part and there's no checker?
if (Checker="" AND Checker2="" AND zcancel="No")
    if GroupMarker <> "P"
        YesNo  "There's no checker. Do you want to fix it?"
        If clipboard()="Yes"
            zcancel="No"
            stop
        endif
    endif
endif

;; Is it missing a puller?
if (Puller = "" AND Puller2 = "" AND zcancel="No")
    if  info("Selected") = info("Records") 
        YesNo  "There's no puller. Do you want to fix it?"
        If clipboard()="Yes"
        zcancel="No"
        	stop
    	endif
	endif
endif

debug

«1SpareText» = datepattern(today(),"mm/dd/yy") + " " + timepattern(now(),"HH:MM AM/PM")

if Status="Com"
	;; if status is already complete, don't change fill date
else
    if info("trigger")="Button.Complete" OR ((OrderNo > 500000 and OrderNo < 600000) and (Order notcontains "ships later" AND Order≠"") AND (Order notcontains "backorder" AND Order≠"")) OR zcancel="Yes"
        ;; stops you from completing orders with backorders or late-shipping items
        if Order contains "backorder" or Order contains "ships later"
            message "This order is not complete"
            zcancel="No"
            stop
        endif
        ;; otherwise sets status as complete
        Status="Com"
        ;;clears out the backorder field if necessary
        if BackOrder≠""
            BackOrder=""
        endif
        ;; sets the filldate
        FillDate=today()
             
        ;; makes sure small Bulbs orders are okay
        if OrderNo >= 500000 and OrderNo < 600000 and form="bulbspagecheck" and zcancel="No"
            if Subtotal<10.00
                YesNo "The total is less than $10.00. Is this ok?"
                if clipboard()="No"
                    Subtotal=10.00
                endif
            endif
        endif
    endif
    
    if info("trigger")="Button.BackOrder" OR ((OrderNo > 500000 and OrderNo < 600000) and (Order contains "ships later" OR Order contains "backorder"))
        if Order CONTAINS "backorder" OR Order CONTAINS "ships later"
            FillDate=today()
            Status="B/O"
            «in_process»=""
    	       call ".backorder"
    	       if info("formname")="mtpagecheck" or info("formname")="bulbspagecheck"
                     «9SpareText»=?(«9SpareText»='','',«9SpareText»+¶+¶)+"---------- ADJUSTED (Status = " + Status +") "+datepattern(today(),"mm/dd/yy")+" ----------"+¶+Order
                    «in_process»=""
                    call "nextorder/1"
                    zcancel="No"
                stop
		          endif
        else
            message "This order has no backorders or stuff to ship later! Check it out."
        endif
	   endif
endif

if GroupMarker contains "G" or zcancel="Yes"
    ;; skip change/h for group cover sheets
    ;; does this result in missing the whole baldue/refund piece?
    ;; don't want to recalculate sales tax when completing groups because it is already done in the Cmd-5 macro, 
    ;; and the TaxedAmount gets screwed up
    is_completing_group = "true"
    ;; maybe call .retotal macro since the MT Totaller adjustment macro is not getting called?
    zcancel="No"
    call ".retotal"
else
	   ;; load the order into the Totaller and run the adjustment macro
	   call "change/h"
endif

;; the MT Totaller adjustment macro calls .retotal, which I think we want to keep... 
;; but it duplicates a lot of the end of the .done macro, so for POE orders, I'm stopping the .done macro right here.
if OrderNo > 600000 and OrderNo < 700000
    zcancel="No"
    stop
endif

debug

;; zero out the Actual_Donated_Refund
Paid = Paid + Actual_Donated_Refund
Actual_Donated_Refund = 0

«BalDue/Refund»=Paid-GrTotal

case «BalDue/Refund»=0 and FillDate=today()
    ;; Donated_Refund=0 ;; never change the Donated_Refund -- this is the MAX donated refund
    Actual_Donated_Refund = 0
case «BalDue/Refund» > 0 and «1stRefund» > 0
	   if «Donated_Refund» ≥ «BalDue/Refund»
    	   «Actual_Donated_Refund» = «BalDue/Refund» 
    	   If «Actual_Donated_Refund»>(.2*OrderTotal) AND «Actual_Donated_Refund»>7
            YesNo "MOFGA is getting more than 20% of this order. OK?"
            if clipboard()="No"
                zcancel="No"
                stop
            endif
        endif
        Message "MOFGA Refund is "+pattern(Actual_Donated_Refund,"$#.##")
    	   Paid=Paid-«BalDue/Refund»
    	   «BalDue/Refund»=0
	   else
    	   If «Actual_Donated_Refund»>(.2*OrderTotal) AND «Actual_Donated_Refund»>7
            YesNo "MOFGA is getting more than 20% of this order. OK?"
            if clipboard()="No"
                zcancel="No"
                stop
            endif
        endif
		      Paid=Paid-«Donated_Refund»
        «Actual_Donated_Refund» = «Donated_Refund»
    	   «BalDue/Refund»=Paid-GrTotal
    	   Message "MOFGA Refund is "+pattern(Actual_Donated_Refund,"$#.##")
    endif
endcase

if «BalDue/Refund» > 0 and (info("trigger")="Button.Complete" OR ((OrderNo > 500000 and OrderNo < 600000) and (Order notcontains "ships later" AND Order≠"") AND (Order notcontains "backorder" AND Order≠"")))
    if Subtotal=0 and «$Shipping»>0 and OrderNo >= 300000 and OrderNo < 400000
    	;; seems like we're intending to not refund shipping on OGS orders?
		call "nextorder/1"
    else
        call ".refund"
    endif
endif

debug

if «BalDue/Refund»<0 and (info("trigger")="Button.Complete" OR ((OrderNo > 500000 and OrderNo < 600000) and (Order notcontains "ships later" AND Order≠"") AND (Order notcontains "backorder" AND Order≠"")))
    call ".baldue"
endif

RealTax=SalesTax

if OrderNo=int(OrderNo)
    Patronage=GrTotal-Donation-CatalogDefrayment-SalesTax
else 
    Patronage=0
endif

if OrderNo >= 500000 and OrderNo < 600000 and ShipCode contains "H"
    message "This order needs to be held. It may be in the holey land. Make sure to match things up."
    «MailFlag»=""
endif

«in_process»=""

call "nextorder/1"


___ ENDPROCEDURE .done _________________________________________________________

___ PROCEDURE .draw_other_orders _______________________________________________
superobject "orderfinder","FillList"
___ ENDPROCEDURE .draw_other_orders ____________________________________________

___ PROCEDURE .email ___________________________________________________________
if Email≠""
    if info("trigger") = "Button.copy email"
        clipboard()="'"+Con+"' <"+Email+">"
    else
        clipboard()=Email
    endif
else
    message "There's no email address. Call or write."
endif
___ ENDPROCEDURE .email ________________________________________________________

___ PROCEDURE .findlost ________________________________________________________
select EntryDate<today()-7 and Order notcontains "backorder" and ShipCode≠"T" and Status≠"Com" and ShipCode≠"Q" and ShipCode≠"J" and (OrderNo<600000)
___ ENDPROCEDURE .findlost _____________________________________________________

___ PROCEDURE .findorders ______________________________________________________
case val(extract(custorder,¬,1)) > 300000 and val(extract(custorder,¬,1)) < 400000
    if info("windows") notcontains "ogspagecheck"
        window "45ogstally:mtpagecheck"
        goform "ogspagecheck" 
    else
        window "45ogstally:ogspagecheck"
    endif
    find OrderNo=val(extract(custorder,¬,1))
    
case val(extract(custorder,¬,1)) > 600000 and val(extract(custorder,¬,1)) < 700000
    if info("windows") notcontains "mtpagecheck"
        window "45ogstally:ogspagecheck"
        goform "mtpagecheck" 
    else
        window "45ogstally:mtpagecheck"
    endif
    find OrderNo=val(extract(custorder,¬,1))
    
case val(extract(custorder,¬,1)) > 500000 and val(extract(custorder,¬,1)) < 600000
        goform "bulbspagecheck" 
    find OrderNo=val(extract(custorder,¬,1))
endcase

cname=""
cgroup=""
define orderlist,?(«C#Test»≠"",lookupalldouble("45orders","C#Text",«C#Text»,"ShipCode", "OrderNo",¶,¬),"Part of a"+¶+" Group")
superobject "orderfinder","filllist"
drawobjects
___ ENDPROCEDURE .findorders ___________________________________________________

___ PROCEDURE .find_early_poe_no_spuds _________________________________________
openfile "SubWizard"
window "SubWizard"
;; select all orders that contain spuds
select (Item >= 7080 and Item <= 7430) or (Item >= 7595 and Item <= 7940)

window "45ogstally"
;; select all orders that contain spuds and then select reverse
select lookupselected("SubWizard","OrderNo",OrderNo,"OrderNo",0,0)
selectreverse

window "SubWizard"
;; select all orders that contain OGS items or onion plants or sweet potatoes
select Item >= 8000 or arraycontains(onionsAndSwPotatoes,str(Item),",")

window "45ogstally"
;; select all orders that contain OGS items or onion plants or sweet potatoes but NOT spuds
selectwithin lookupselected("SubWizard","OrderNo",OrderNo,"OrderNo",0,0) and (Order contains '7990' or Order contains '7995')
;; selectwithin lookupselected("SubWizard","OrderNo",OrderNo,"OrderNo",0,0)

if info("empty")
    stop
else
    message "okay, print these with the regular picksheet print macro"
endif

;; now we have a selection of orders that contain ginger and/or turmeric, and SOMETHING else, but not spuds
;; so these orders should be printed using the regular picksheet print macro and form

___ ENDPROCEDURE .find_early_poe_no_spuds ______________________________________

___ PROCEDURE .find_uncompleted_onions _________________________________________
arrayfilter onionsAndSwPotatoes,onionsAndSwPotatoes,",","import() contains "+quoted(import())
onionsAndSwPotatoes=replace(onionsAndSwPotatoes,","," or ")

select Status notcontains "c" and OrderNo > 600000 and OrderNo < 700000 and «BalDue/Refund» = 0 and GroupMarker = ''
selectwithin arraystrip(arrayfilter(replace(Order,lf(),cr()),cr(),{?(}+onionsAndSwPotatoes+{,"",import())}),cr())=""

if info("empty")
    message "all dropship-only orders are completed"
endif
___ ENDPROCEDURE .find_uncompleted_onions ______________________________________

___ PROCEDURE .grp_count _______________________________________________________
arraybuild zprintrun, ",", "",str(int(OrderNo)) ; build a list of all selected order numbers
arrayfilter zprintrun, zorderarray, ",", ?(import()=str(int(OrderNo)),str(int(OrderNo)),"") ;strip to the order being worked on
zorderarray=arraystrip(zorderarray, ",") ;strip out blank elements
zordercount=arraysize(zorderarray,",")
«1SpareNumber»=zordercount-1
;message zordercount
___ ENDPROCEDURE .grp_count ____________________________________________________

___ PROCEDURE .holdorder _______________________________________________________
local yo, orig_comments, orig_notes
yo = datepattern(«DeliveryDate»,"mm/dd/yyyy")
gettext "When should this order be held until?", yo
«DeliveryDate»=date(yo)
if yo <> ""
    orig_comments = OrderComments
    orig_notes=«Notes4»
    «Notes4»=orig_notes + "Scheduled date is " + yo
    OrderComments = orig_comments + "---original shipcode: " + ShipCode
    ShipCode = "Q"
endif
___ ENDPROCEDURE .holdorder ____________________________________________________

___ PROCEDURE .importorders ____________________________________________________
local filem, folderm, typem
ono=OrderNo

waswindow = info("windowname")

openfiledialog folderm, filem, typem, "ZEPD"
if filem=""
    stop
endif

openfile "++" + folderpath(folderm)+filem

Field "C#Text"
Select «C#Text» = ""
FormulaFill str(«C#»)
SelectAll
field OrderNo
sortup
find OrderNo=ono


if waswindow contains "mtpagecheck" or waswindow contains "bulbspagecheck"
    message "remember to populate OGSOrder"
endif
___ ENDPROCEDURE .importorders _________________________________________________

___ PROCEDURE .indigenous_royalties ____________________________________________
global order, neworder
local order_len
local allwindows, numwindows, n, onewindowname

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; This macro creates a vertical file of all OGS, POE, and Bulbs (if any)
;; items on which we pay Indigenous Royalties. It is intended to be run
;; at the end of the fiscal year after all orders have been completed.
;; Requires an unlinked file, 45indigenousroyalties.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

waswindow=info("windowname")
 
;;==========================================

select  ShipCode notcontains "D" and OrderNo > 0 and Status <> "" and GroupMarker notcontains "G" and Order contains ")"

field OrderNo
sortup

openfile "45indigenousroyalties"
DeleteAll

window waswindow
neworder=""
order=""
firstrecord

debug

;;NoShow
loop
    order=Order
    order_len = arraysize(extract(order,¶,1),¬)
    
    case (order_len = 10 or order_len = 11) and OrderNo < 400000 ;; OGS
        ArrayFilter order,neworder,¶,ShipCode+¬+str(OrderNo)+¬+str(Sequence)+¬+import()
    case order_len = 8 and OrderNo > 500000 and OrderNo < 600000 ;; Bulbs
        ArrayFilter order,neworder,¶,ShipCode+¬+str(OrderNo)+¬+str(Sequence)+¬+extract(import(),¬,1)+¬+extract(import(),¬,2)+¬+¬+extract(import(),¬,3)+¬+extract(import(),¬,4)+¬+extract(import(),¬,5)+¬+extract(import(),¬,6)+¬+extract(import(),¬,7)+¬+¬+extract(import(),¬,8)
    case order_len = 11 and OrderNo > 600000 ;; POE
        ArrayFilter order,neworder,¶,ShipCode+¬+str(OrderNo)+¬+str(Sequence)+¬+extract(import(),¬,1)+¬+extract(import(),¬,2)+"-"+extract(import(),¬,3)+¬+str(val(extract(import(),¬,4)))+¬+extract(import(),¬,5)+¬+extract(import(),¬,6)+¬+extract(import(),¬,7)+¬+extract(import(),¬,8)+¬+extract(import(),¬,9)+¬+extract(import(),¬,10)+¬+extract(import(),¬,11)+¬+extract(import(),¬,12)
    case order_len = 9 and OrderNo > 600000 ;; POE unprinted order (onions or sweet potatoes)
        ;; do nothing
    defaultcase
        message "something weird in the Order field for " + str(OrderNo) + "!" + " Array length is " + str(order_len) + "."
        stop
    endcase

    window "45indigenousroyalties:secret"
    openfile "+@neworder"
    window waswindow
    downrecord
    order=""
    neworder=""
until info("stopped")
;;EndNoShow

;;========================================================

window "45indigenousroyalties"

Field qty
Formulafill ?(comment contains "out-of-stock",0,qty)

selectall

select Item <> "" and Item <> "0"
removeunselected

;; get this list from the catalog or from the ogs_cat, moose_cat, and bulbs_cat tables on the website database (in the key fields)
select Item = 8058 or Item = 8615 or Item = 7228 or Item = 7230 or Item = 7240 or Item = 7245 or Item = 7259 or Item = 7270 or Item = 7640 or Item = 7735 or Item = 7740 or Item = 7745 or Item = 7750 or Item = 7765 or Item = 7790 or Item = 7800 or Item = 7875 or Item = 7900 or Item = 7905 or Item = 7910
removeunselected

select comment notcontains "out-of-stock"
removeunselected

field discount_rate
formulafill lookup("45ogstally","OrderNo",OrderNo,"Discount",0,0)

field amount_paid
formulafill «$Total»*(1-discount_rate)

Save
___ ENDPROCEDURE .indigenous_royalties _________________________________________

___ PROCEDURE .just_hold_order _________________________________________________
if ShipCode≠"H"
    Notes1=?(Notes1="","ShipCode started "+ShipCode+" - "+datepattern(today(),"mm/dd"),
                Notes1+¶+"ShipCode started as "+ShipCode+" - "+datepattern(today(),"mm/dd"))
    ShipCode="H"
else
    call ChangeShipCode
endif
___ ENDPROCEDURE .just_hold_order ______________________________________________

___ PROCEDURE .lookatorders ____________________________________________________
global offorder, offaddress
offorder=""
offaddress=""
waswindow=info("windowname")
case  val(extract(findinorder,¬,2))>500000 and val(extract(findinorder,¬,2))<600000
    goform "bulbspagecheck"
    selectall                          
    find OrderNo=val(extract(findinorder,¬,2))

    superobject "orderfinder", "FillList"
    drawobjects
case (OrderNo>300000 and OrderNo<400000)
    goform "ogspagecheck"
    selectall                          
    find OrderNo=val(extract(findinorder,¬,2))

    superobject "orderfinder", "FillList"
    drawobjects
case OrderNo≥600000 and OrderNo<700000
    goform "mtpagecheck"
    selectall                          
    find OrderNo=val(extract(findinorder,¬,2))

    superobject "orderfinder", "FillList"
    drawobjects
endcase                        
___ ENDPROCEDURE .lookatorders _________________________________________________

___ PROCEDURE .moose_pool_selection ____________________________________________
global mpool, msizegroup, mshipcodetosearch, mstates, usps, batch_num
local howmanyshipcodes

if info("trigger")="Button.Reset"
    mpool = "1"
    msizegroup = "a"
    mshipcodetosearch = "U,X"
    drawobjects
    stop
endif

if info("trigger")="Button.Search"
;; find the largest batch number so we don't reuse numbers
selectall
field «1SpareNumber»
maximum
batch_num = «1SpareNumber» + 1
deleterecord

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;; POOL ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
case mpool="1"
    select arraycontains("AL,AR,AZ,CA,DC,FL,GA,HI,LA,MD,MS,NC,OR,SC,TN,TX,VA,WA",Sta,",")
    mstates = "Alabama, Arkansas, Arizona, California, DC, Florida, Georgia, Hawaii, Louisiana, Maryland, Mississippi, North Carolina, Oregon, South Carolina, Tennessee, Texas, Virginia, Washington"
case mpool="2"
    select arraycontains("DE,IL,KS,MO,NJ,NM,NV,OK,UT",Sta,",")
    mstates = "Delaware, Illinois, Kansas, Missouri, New Jersey, New Mexico, Nevada, Oklahoma, Utah"
case mpool="3"
    select arraycontains("IN,KY,OH,WV",Sta,",")
    mstates = "Indiana, Kentucky, Ohio, West Virginia"
case mpool="4"
    select arraycontains("CT,PA",Sta,",")
    mstates = "Connecticut, Pennsylvania"
case mpool="5"
    select arraycontains("MA,RI",Sta,",")
    mstates = "Massachusetts, Rhode Island"
case mpool="6"
    select arraycontains("CO,IA,ID,NE,WI",Sta,",")
    mstates = "Colorado, Iowa, Idaho, Nebraska, Wisconsin"
case mpool="7"
    select arraycontains("MI",Sta,",")
    mstates = "Michigan"
case mpool="8"
    select arraycontains("NY",Sta,",")
    mstates = "New York"
case mpool="9"
    select arraycontains("NH",Sta,",")
    mstates = "New Hampshire"
case mpool="10"
    select arraycontains("MN,MT,ND,SD,WY",Sta,",")
    mstates = "Minnesota, Montana, North Dakota, South Dakota, Wyoming"
case mpool="11"
    select arraycontains("VT",Sta,",")
    mstates = "Vermont"
case mpool="12"
    select arraycontains("ME",Sta,",")
    mstates = "Maine"
case mpool="13"
    select arraycontains("AK",Sta,",")
    mstates = "Alaska"
endcase


;;;;;;;;;;;;;;;;;;;;;;;;;; SHIP CODES ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
howmanyshipcodes = arraysize(mshipcodetosearch,",")
local n
global shipcodesearchstring

shipcodesearchstring = "SelectWithin"
n = 1

loop
    shipcodesearchstring = shipcodesearchstring + " ShipCode contains '" + array(mshipcodetosearch,n,",") + "' OR "
    n = n+1
while n <= howmanyshipcodes

if howmanyshipcodes > 0
    shipcodesearchstring = trim(shipcodesearchstring,4)
    execute shipcodesearchstring
    if info("empty")
        message "no records found"
        stop
    endif
endif

;;;;;;;;;;;;;;;;;;;;;;;;;;;;; PACKAGE SIZES ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
case msizegroup="s"
    selectwithin OrderNo >= 600000 and PickSheet notcontains "1" and PickSheet <> "0" and (ShipCode contains "U" or ShipCode contains "X") and Status notcontains "c" and Order notcontains "0000–" and Order notcontains "1)	0–" and GroupMarker notcontains "P" and GroupMarker notcontains "G" ;; and Sequence < 6720
    if info("empty")
        message "no records found"
        stop
    endif
    selectwithin Order notcontains "–C" and Order notcontains "–D" and Order notcontains "–E" and Order notcontains "(50)" and Order notcontains "(25)" and Order notcontains "(10)" and Order notcontains "50 lbs" and Order notcontains "25 lbs" and Order notcontains "10 lbs"
    if info("empty")
        message "no records found"
        stop
    endif
case msizegroup="l"
    selectwithin OrderNo >= 600000 and PickSheet notcontains "1" and PickSheet <> "0" and (ShipCode contains "U" or ShipCode contains "X") and Status notcontains "c" and Order notcontains "0000–" and Order notcontains "1)	0–" and GroupMarker notcontains "P" and GroupMarker notcontains "G" ;; and Sequence < 6720
    if info("empty")
        message "no records found"
        stop
    endif
    selectwithin Order notcontains "–B" and Order notcontains "–C" and Order notcontains "–D" and Order notcontains "(2)" and Order notcontains "(0.5)" and Order notcontains "(10)" and Order notcontains "(6)" and Order notcontains "6 lbs" and Order notcontains "2 lbs" and Order notcontains "10 lbs" and ((Order notcontains "(25)" and Order notcontains "25 lbs") or Order contains "7400–D" or Order contains "7400	D") 
    if info("empty")
        message "no records found"
        stop
    endif
case msizegroup="a"
    selectwithin OrderNo >= 600000  and PickSheet notcontains "1" and PickSheet <> "0" and Status notcontains "c" and GroupMarker notcontains "P" and GroupMarker notcontains "G" ;; and Sequence < 6720
    if info("empty")
        message "no records found"
        stop
    endif
case msizegroup="g"
    selectwithin OrderNo >= 600000 and PickSheet notcontains "1" and PickSheet <> "0" and Status notcontains "c" and (GroupMarker contains "P" or GroupMarker contains "G") and Sequence < 6720
    create_batches = "false"
    if info("empty")
        message "no records found"
        stop
    endif
endcase


endif
;selectwithin OrderPlaced <= date("3/10/21")
if info("empty")
    message "no more orders"
    selectall
endif

debug

if info("trigger")="Button.Print Selection with Cover"
    Field Sequence
    Sortup

    ;; print cover sheet
    openform "mt_batch_coversheet"
    printonerecord dialog
    CloseWindow
    
    ;; print picksheets: bypass PicksheetPrint macro
    global oldwindow, oldfile, pickno

    oldfile=info("DatabaseName") 
    oldwindow=info("windowname")
    waswindow="mtpagecheck"
    call ".mtpicksheets"

endif
___ ENDPROCEDURE .moose_pool_selection _________________________________________

___ PROCEDURE .mtearlypicksheets _______________________________________________
global create_batches, msizegroup, batch_num, is_early

create_batches = "false"

pickno=0
waswindow=info("windowname")

if info("files") notcontains "MT Totaller"
    openfile "MT Totaller"
endif
    
if info("files") notcontains "SubWizard"
    openfile "SubWizard"
   synchronize
else 
    window "SubWizard"
    synchronize
endif

if info("files") notcontains "45mt prices"
    openfile "45mt prices"
endif

window waswindow
NoYes "run selection?"
If clipboard()="Yes"
    goto ready
endif

NoYes "do you want to do zipcodes?"
if clipboard()="Yes"
    call ".zippick"
    goto ready
endif

ready:
YesNo "print picksheets in batches?"
if clipboard()="Yes"
    create_batches = "true"
endif

selectwithin OrderNo>600000
selectwithin PickSheet="" and Status="" and AdditionalPickSheet=""
selectwithin extract(extract(Order,¶,1),¬,2)≠"0000"
selectwithin (Order contains "7990" or Order contains "7995")

if info("empty")
    stop
endif

field Sequence
sortup
call ".commentsearlymt"

if create_batches = "true"
    msizegroup = "e"
    getscrap "Set the batch number"
    batch_num = val(clipboard())
    call .create_poe_batches
    is_early = "true"
    
    debug
    
    firstrecord
    loop
        call .print_poe_batches
        downrecord
    until info("eof")
endif

openform "mtearlypicksheet"
print dialog
CloseWindow

window waswindow
removesummaries 7

if needquotes <> ""
    bigmessage "These orders need shipping quotes: " + needquotes
endif
___ ENDPROCEDURE .mtearlypicksheets ____________________________________________

___ PROCEDURE .mtpicksheets ____________________________________________________
global vedad, is_early
pickno=0
is_early = "false"
needquotes=""


waswindow=info("windowname")

;;goto skippy

if info("files") notcontains "MT Totaller"
    openfile "MT Totaller"
    endif
if info("files") notcontains "SubWizard"
    openfile "SubWizard"
   synchronize
else 
    window "SubWizard"
    synchronize
endif
if info("files") notcontains "45mt prices"
    openfile "45mt prices"
endif
window waswindow
NoYes "run selection?"
If clipboard()="Yes"
    goto ready
endif
NoYes "do you want to do zipcodes?"
if clipboard()="Yes"
    call ".zippick"
    goto ready
endif

getscrap "first sequence number"
Numb=val(clipboard())
pickno=Numb
select «Sequence»≥Numb 
getscrap "last sequence number"
Numb=val(clipboard())
ono=Numb
selectwithin «Sequence»<= Numb
NoYes "Do you want to choose shipcodes?"
If clipboard()="Yes"
    getscrap "which one?"
    selectwithin ShipCode=upper(clipboard())
    loop
        NoYes "Another"
        stoploopif clipboard()="No"
        getscrap "which one?"
        selectadditional «Sequence»≥pickno And «Sequence»<=ono And ShipCode=upper(clipboard())
    while forever
endif


ready:


selectwithin OrderNo>600000
selectwithin PickSheet notcontains "1" and Status<>"Com"
selectwithin extract(extract(Order,¶,1),¬,2)≠"0000"

if info("empty")
    stop
endif

debug



;;field Sequence
;;sortup
call ".commentsmt"

skippy:

if create_batches = "true"
    call .create_poe_batches
        
    firstrecord
    loop
        call .print_poe_batches
        downrecord
    until info("eof")
endif

openform "mtpicksheet"
print dialog
CloseWindow

;;selectwithin
;;arraysize(Order,¶)>45

;;if info("empty")
;;    goto done
;;else
;;    openform "mtpicksheet2"
;;    print dialog
;;    CloseWindow
;;endif

;;done:
window waswindow

removesummaries 7

if needquotes <> ""
    displaydata "These orders need shipping quotes: " + arraystrip(needquotes,¶)
endif
___ ENDPROCEDURE .mtpicksheets _________________________________________________

___ PROCEDURE .mtreprint _______________________________________________________
openfile "MT Totaller"
oldwindow=info("windowname")
window waswindow
order=Order
order=replace(order, "(","")
order=replace(order,")","")
vship=ShippingWt
f=«Bulk»
window "MT Totaller:secret"
openfile "&@order"
call ".adjustment"
openform "mtpicksheet"
printonerecord dialog
CloseWindow

if arraysize(rayb,¶)<50
    goto done
else
    openform "mtpicksheet2"
    printonerecord dialog
    CloseWindow
endif

done:
window waswindow
___ ENDPROCEDURE .mtreprint ____________________________________________________

___ PROCEDURE .ogspicksheet ____________________________________________________
global vedad, batch_num, msizegroup

create_batches = "false"

selectall

waswindow=info("windowname")
getscrap "First order to print:"
Numb=val(clipboard())
Numb=?(Numb<300000,Numb+300000,Numb)
pickno=Numb
getscrap "Last order:"
secno=val(clipboard())
secno=?(secno<300000,secno+300000,secno)
ono=Numb
select OrderNo≥Numb and OrderNo < (secno+1)
if info("empty")
    message "Orders " + str(Numb) + " - " + str(secno) + " are not in the tally yet!"
    stop
endif

if info("selected")=info("records")
    message "Something is screwy"
    stop
endif

;; uncomment these lines when we are ready to try batches in OGS

;;YesNo "print picksheets in batches?"
;;if clipboard()="Yes"
;;    create_batches = "true"
;;endif

;;updates discount table
yesno "Update Discount Table?"
case clipboard() contains "yes"
    window "discounttable"
    selectall
    window waswindow
    local quien,quanto,quantogrande,nombre,groupe,emale,etat,mass
    firstrecord
    loop
        quien=«C#»
        quanto=AdjTotal
        quantogrande=Subtotal
        nombre=Con
        groupe=Group
        emale=Email
        etat=St
        if Bulk="bulk"
            mass=1
        else
            mass=0
        endif
        window "discounttable"
        find «C#»=quien
        if info("found")
            TallyPurchases=TallyPurchases+quanto
            thisyeartallypurchases=thisyeartallypurchases+quanto
            window waswindow
            downrecord
        else
            if quantogrande≥100
                insertrecord
                «C#»=quien
                Con=nombre
                Group=groupe
                email=emale
                State=etat
                if mass=1
                    Bulk=1
                else
                    Bulk=0
                endif
                TallyPurchases=TallyPurchases+quanto
                thisyeartallypurchases=thisyeartallypurchases+quanto
                DateAdded=today()
                window waswindow
                downrecord
            else
                window waswindow
                downrecord
            endif
        endif
        if OrderNo≠int(OrderNo)
            loop
            downrecord
            until OrderNo=int(OrderNo) or info("stopped")
        endif
    until info("stopped")
    window "discounttable"
    select TotalPurchases≠WalkInPurchases+TallyPurchases or DateAdded=today()
    if info("selected")=info("records")
        window waswindow
    else
        field TotalPurchases
        call tabdown
        call discountfill
        selectall
        window waswindow
    endif
case clipboard() contains "No"
    goto scrapthediscountbit
endcase

scrapthediscountbit:
firstrecord
loop
«C#Text»=str(«C#»)
if ShipCode="U" and ((Zip<9999 and (chunkarray(Str(Zip),1,2)="12" or chunkarray(Str(Zip),1,2)="25" or chunkarray(Str(Zip),1,2)="28" or chunkarray(Str(Zip),1,2)="29" or chunkarray(Str(Zip),1,2)="32" 
or chunkarray(Str(Zip),1,2)="39" or chunkarray(Str(Zip),1,2)="40" or chunkarray(Str(Zip),1,2)="43" or chunkarray(Str(Zip),1,2)="43" or chunkarray(Str(Zip),1,2)="44"
or chunkarray(Str(Zip),1,2)="46" or chunkarray(Str(Zip),1,2)="47" or chunkarray(Str(Zip),1,2)="52" or chunkarray(Str(Zip),1,2)="53" or chunkarray(Str(Zip),1,2)="58" or chunkarray(Str(Zip),1,2)="60"
or chunkarray(Str(Zip),1,2)="85" or chunkarray(Str(Zip),1,2)="88")) or (Zip>9999 and (chunkarray(Str(Zip),1,3)="120" or chunkarray(Str(Zip),1,3)="125" or chunkarray(Str(Zip),1,3)="131"
or chunkarray(Str(Zip),1,3)="138" or chunkarray(Str(Zip),1,3)="144" or chunkarray(Str(Zip),1,3)="148" or chunkarray(Str(Zip),1,3)="195")))
    «3SpareText»="dpt?"+?(Zip<9999, "0"+chunkarray(Str(Zip),1,2), chunkarray(Str(Zip),1,3))
endif
if Order contains "pallet"
«5SpareText»="PULLERS! Order contains pallet quantity. Double-check bag count."
endif
if ShipCode="C"
«ShipmentDescription»="Farm Supplies"
endif
if OrderPlaced=0
OrderPlaced=today()
endif
if GroupMarker notcontains "P" and GroupMarker notcontains "G"
    local tabledisc
    tabledisc=lookup("discounttable","C#",«C#»,"Discount",0,0)
    if tabledisc>Discount and Bulk notcontains "bulk" and Bulk notcontains "staff"
        Discount=tabledisc
        Notes3 = Notes3 + " --- rolling discount caused discount rate to go up"
    endif
endif
if OrderNo≠int(OrderNo)
field «C#»
cutcell
endif
downrecord
until info("stopped")

selectwithin PickSheet=""
selectwithin ShipCode notcontains "D"
selectwithin extract(extract(Order,¶,1),¬,2)≠"0000"
if info("empty")
    stop
endif
;goes through .commentsogs procedure

global order, waswindow, order_no
openfile "NewOGSTotaller"
window "45ogscomments.linked"
Synchronize
window "45ogscomments"
openfile "&&" + "45ogscomments.linked"

window waswindow
firstrecord
loop
    order_no=OrderNo
    order=Order
    Sh=ShipCode
    vedad=Sta
    window "NewOGSTotaller"
    openfile "&@order"
    call ".commentfill"
    downrecord
until info("stopped")


if create_batches = "true"
    msizegroup = "o"
    getscrap "Set the batch number"
    batch_num = val(clipboard())
        
    call .create_ogs_batches
    
    debug
        
    firstrecord
    loop
        call .print_ogs_batches
        downrecord
    until info("eof")
endif

find «1SpareNumber» = 999999 and info("summary") = 1
if info("found")
    deleterecord
endif

find info("summary") = 2
if info("found")
    deleterecord
endif

debug

openform "picksheetOGS"
print dialog
CloseWindow

/*--------------new: causes next pages of picksheet to print------------*/
openform "picksheet2OGS"
firstrecord
loop
    if arraysize(PickSheet,¶)>30 and «1SpareNumber» <> 999999
        printonerecord ""
    endif
    downrecord
until info("stopped")
CloseWindow
openform "picksheet3OGS"
firstrecord
loop
    if arraysize(PickSheet,¶)>85 and «1SpareNumber» <> 999999
        printonerecord ""
    endif
    downrecord
until info("stopped")
CloseWindow
/*---------------------------------------------------------------------------------*/

selectwithin (ShipCode = "P" or ShipCode = "T") and OrderNo = int(OrderNo)
if info("empty")
else
    openform "Pickup label"
    print dialog
    CloseWindow
endif

window waswindow
removesummaries 7
selectall
field OrderNo
sortup
find OrderNo=pickno
save

if needquotes <> ""
    displaydata "These orders need shipping quotes: " + arraystrip(needquotes,¶)
endif
___ ENDPROCEDURE .ogspicksheet _________________________________________________

___ PROCEDURE .ogsreprint ______________________________________________________
waswindow=info("windowname")
openfile "NewOGSTotaller"
oldwindow=info("windowname")
window waswindow
order=Order
window "NewOGSTotaller:secret"
openfile "&@order"
call ".adjustment"
openform "picksheetOGS"
printonerecord dialog
CloseWindow

if arraysize(rayb,¶)<31
    goto done
else
    openform "picksheet2OGS"
    printonerecord dialog
    CloseWindow
endif

if arraysize(rayb,¶)<86
    goto done
else
    openform "picksheet3OGS"
    printonerecord dialog
    CloseWindow
endif

done:
window waswindow
___ ENDPROCEDURE .ogsreprint ___________________________________________________

___ PROCEDURE .onions_exported_today ___________________________________________
«2SpareDate»=today()
___ ENDPROCEDURE .onions_exported_today ________________________________________

___ PROCEDURE .open_order_history_form _________________________________________
local newWindowRect
newWindowRect=rectanglecenter(
    info("screenrectangle"),
    rectanglesize(1,1,9*72,10*72))
setwindowrectangle newWindowRect,
    ""

openform "history of order changes"
___ ENDPROCEDURE .open_order_history_form ______________________________________

___ PROCEDURE .pickedup ________________________________________________________
DateShipped="Picked up "+datepattern(today(),"mm/dd/yy")
PickedUp="Yes"
___ ENDPROCEDURE .pickedup _____________________________________________________

___ PROCEDURE .poe_populate_ogs_order __________________________________________
local istruck
global vedad

debug

waswindow = info("windowname")

openfile "45ogscomments.linked"
openfile "45ogscomments"
openfile "&&" + "45ogscomments.linked"

    ;; find orders with OGS items and add them to OGSOrder field
    openfile "SubWizard"
    select Item >= 8000
    window waswindow
    select lookupselected("SubWizard","OrderNo",OrderNo,"OrderNo",0,0)
    selectwithin OGSOrder = ''
    
    if info("empty")
        stop
    endif
    
    firstrecord
    openfile "MT Totaller"
    loop
        window waswindow
        istruck=«8SpareText»
        adjusting_order = 1
        vedad=TaxState

        
        if extract(extract(Order,¶,1),¬,3) contains "1" or 
        extract(extract(Order,¶,1),¬,3) contains "2" or
        extract(extract(Order,¶,1),¬,3) contains "3" or
        extract(extract(Order,¶,1),¬,3) contains "4" or
        extract(extract(Order,¶,1),¬,3) contains "5" or
        extract(extract(Order,¶,1),¬,3) contains "6" or
        extract(extract(Order,¶,1),¬,3) contains "7" or
        extract(extract(Order,¶,1),¬,3) contains "8" or
        extract(extract(Order,¶,1),¬,3) contains "9"
            f=Bulk
            order=Order
            order=replace(order, " lbs.","")
            order=replace(order, "(","")
            order=replace(order,")","")
            order=replace(order,"–",¬) ;;//This replaces the n-dash which is what the office uses to connect Item with size.
            order=replace(order,"-",¬) ;;//This replaces a regular dash should you happen to use that.
            arrayfilter order, order, ¶, arrayinsert(extract(order,¶,seq()),4,1,¬)
            window "MT Totaller"
            openfile "&@order"
            call ".newadjustment"
            order = ""
        else
            window waswindow
            order=Order
            order=replace(order, "(","")
            order=replace(order,")","")
            order=replace(order," lbs.","")
            vship=ShippingWt
            f=Bulk
            truck=ShipCode
            istruck =val(«8SpareText»)
            window "MT Totaller"
            openfile "&@order"
            if istruck = 1
                call ".truckadjustment"
            else
                call ".adjustment"
            endif
        
            order = ""
        endif
        
    window waswindow 
    downrecord
    until info("stopped")
   
   message "done"
___ ENDPROCEDURE .poe_populate_ogs_order _______________________________________

___ PROCEDURE .previewsubs _____________________________________________________
local tallyorderno, newWindowRect
tallyorderno = OrderNo
openfile "SubWizard"
select OrderNo = tallyorderno

openfile "SubWizard"
newWindowRect=rectangle(30,20,400,1200)
setwindowrectangle newWindowRect,""
openform "PreviewSubs"
firstrecord
___ ENDPROCEDURE .previewsubs __________________________________________________

___ PROCEDURE .ProduceCSVForCCSchEmail _________________________________________
local csvarray

noyes "Use current selection?"

if clipboard() contains "No" 
    select ShipCode contains "C" and OrderNo = int(OrderNo) and ShippingWt > 0 and OrderNo > 600000 and Notes4 notcontains "exported CSV for scheduling email"
    if info("empty")
        message "There are apparently no new POE orders in need of scheduling emails."
    endif
endif

if info("selected") > 350 or info("empty")
    message "This seems like too many records. Try again."
    stop
endif

arrayselectedbuild csvarray, ¶, "",Con+","+Group+","+Email+","+str(OrderNo)+","+SAd+","+Cit+","+Sta+","+''+","+'1'
clipboard() = csvarray

field Notes4
formulafill Notes4 + ' --- exported CSV for scheduling email ' + datepattern(today(),"mm/dd/yy")

message "Got the email addresses on the clipboard! Paste them into the batch email tool."

shellopendocument "https://staging.fedcoseeds.com/manage_site/send-common-carrier-quote-emails/moose"

___ ENDPROCEDURE .ProduceCSVForCCSchEmail ______________________________________

___ PROCEDURE .ProduceCSVForDepotEmail _________________________________________
local csvarray, backorders_list, which_depot, order

csvarray = ""
order = ""


;; based on which pagecheck form you're on, make sure the selection of orders doesn't include
;; any from other branches

case info("formname")="ogspagecheck"
    Find OrderNo > 400000
case info("formname")="mtpagecheck"
    Find OrderNo < 600000
endcase

if info("found")
    message "Check selection, and remove orders from other branches"
    ;;stop
endif

Find Email='' or ShipCode notcontains "J" or OrderNo <> int(OrderNo) or Status = "" or Depot = ""
if info("found")
    message "Check selection, and remove orders without emails, non-depot orders, uncompleted orders, and group parts."
    ;; stop
endif

firstrecord
which_depot = Depot

loop
    backorders_list = ""
    
    if Depot <> which_depot
        message "This selection seems to contain more than one depot. Adjust selection and try again."
        stop
    endif
    
    if GroupMarker="G" and Status contains "b"
        order = strip(arraystrip(Order,¶))
        arrayfilter order, backorders_list, ¶, extract(extract(order,¶,seq()),¬,1)+" " +extract(extract(order,¶,seq()),¬,2)+" " + extract(extract(order,¶,seq()),¬,9) + "# "+extract(extract(order,¶,seq()),¬,10)
    else    
        arrayfilter arraystrip(BackOrder,¶), backorders_list, ¶, extract(extract(BackOrder,¶,seq()),¬,2)+" " + extract(extract(BackOrder,¶,seq()),¬,9) + "# "+extract(extract(BackOrder,¶,seq()),¬,10)
    endif
    
    backorders_list = replace(backorders_list,¶,";")

    csvarray = csvarray + ¶ + Con+","+Email+","+str(OrderNo)+","+SAd+","+Cit+","+Sta+","+pattern(Z,"#####")+","+backorders_list
    
    downrecord
until info("stopped")

arraystrip csvarray, ¶

clipboard() = csvarray

;field Notes4
;formulafill Notes4 + ' --- sent depot pickup notification email ' + datepattern(today(),"mm/dd/yy")

message "Got the email addresses on the clipboard! Paste them into the bulk emailer."

order = ""

case info("formname")="ogspagecheck"
    shellopendocument "https://fedcoseeds.com/manage_site/send-depot-batch-emails/choose-depot/ogs"
case info("formname")="mtpagecheck"
    shellopendocument "https://fedcoseeds.com/manage_site/send-depot-batch-emails/choose-depot/moose"
endcase


___ ENDPROCEDURE .ProduceCSVForDepotEmail ______________________________________

___ PROCEDURE .ProduceCSVForPickupEmail ________________________________________
local csvarray


;; based on which pagecheck form you're on, make sure the selection of orders doesn't include
;; any from other branches

case info("formname")="ogspagecheck"
    Find OrderNo > 400000
case info("formname")="mtpagecheck"
    Find OrderNo < 600000
case info("formname")="bulbspagecheck"
    Find OrderNo < 500000 or OrderNo > 600000
endcase

if info("found")
    message "Check selection, and remove orders from other branches"
    stop
endif

Find Email='' or ShipCode notcontains "P" or OrderNo <> int(OrderNo) or PickedUp <> '' or Status notcontains "c"
if info("found")
    message "Check selection, and remove orders without emails, non-pickup orders, uncompleted orders, orders already picked up, and group parts."
    stop
endif

arrayselectedbuild csvarray, ¶, "",Con+","+Email+","+str(OrderNo)+","+datepattern(PickUpDate,"mm/dd/yy")

clipboard() = csvarray

field Notes4
formulafill Notes4 + ' --- sent pickup notification email ' + datepattern(today(),"mm/dd/yy")

message "Got the email addresses on the clipboard! Paste them into the bulk emailer."

case info("formname")="ogspagecheck"
    shellopendocument "https://fedcoseeds.com/manage_site/send-pickup-batch-emails/ogs"
case info("formname")="mtpagecheck"
    shellopendocument "https://fedcoseeds.com/manage_site/send-pickup-batch-emails/moose"
case info("formname")="bulbspagecheck"
    shellopendocument "https://fedcoseeds.com/manage_site/send-pickup-batch-emails/bulbs"
endcase


___ ENDPROCEDURE .ProduceCSVForPickupEmail _____________________________________

___ PROCEDURE .ProduceCSVForShQuoEmail _________________________________________
local csvarray

noyes "Use this selection?"

if clipboard() contains "Yes"
    arrayselectedbuild csvarray, ¶, "",Con+¬+Group+¬+Email+¬+str(OrderNo)+¬+SAd+¬+Cit+¬+Sta+¬+pattern(Z,'#####')+¬+str(ShippingWt)
    clipboard() = csvarray

    field Notes4
    formulafill Notes4 + ' --- exported CSV for ship quote ' + datepattern(today(),"mm/dd/yy")
else
    select ShipCode notcontains "P" and OrderNo = int(OrderNo) and «$Shipping» = 0 and ShippingWt > 0 and OrderNo > 600000 and Depot notcontains "NOFA" and Notes4 notcontains "exported CSV for ship"
    if info("empty")
        message "There are apparently no new orders in need of shipping quotes."
    endif
    ;; this check is not actually necessary if running the search above
    Find OrderNo <> int(OrderNo) or «$Shipping» <> 0 or ShippingWt = 0 or ShipCode contains "P" or Notes4 contains "exported CSV for ship quote"
    if info("found")
        message "Check selection, and remove group parts, orders that already have shipping charges, P shipcodes, orders that have already been exported to CSV."
        stop
    endif

    arrayselectedbuild csvarray, ¶, "",Con+¬+Group+¬+Email+¬+str(OrderNo)+¬+SAd+¬+Cit+¬+Sta+¬+pattern(Z,'#####')+¬+str(ShippingWt)

    clipboard() = csvarray

    field Notes4
    formulafill Notes4 + ' --- exported CSV for ship quote ' + datepattern(today(),"mm/dd/yy")
endif

message "Got the email addresses on the clipboard! Paste them into the Google spreadsheet."

shellopendocument "https://docs.google.com/spreadsheets/d/1nC8mpKcpXQ-zVs5tJRYC5QFWOI4AqvZ4WWgVPH8RPpM/edit?usp=sharing"


___ ENDPROCEDURE .ProduceCSVForShQuoEmail ______________________________________

___ PROCEDURE .print_ogs_batches _______________________________________________
global collector, order_aggregate
order_aggregate = ""
collector = ""

window waswindow
field «1SpareNumber»
groupup
field «2SpareNumber»
propagate

Field Order
FormulaFill ?(info("summary") =0,«»,collector+assign((sandwich("",collector,¶)+«»)[1,info("summary") =0],"collector"))

;; just show summary records
outlinelevel "1"

Field Con
FormulaFill str(«2SpareNumber») + " Batch " + str(«1SpareNumber») + " Summary"

;; run the summary records through a special macro in the totaller
firstrecord
loop
    order_aggregate = Order
    order_aggregate=replace(order_aggregate, "(","")
    order_aggregate=replace(order_aggregate,")","")
    order_aggregate=replace(order_aggregate," lbs.","")
    openfile "NewOGSTotaller"
    openfile "&@order_aggregate"
    call .create_batch_summaries
    downrecord
until info("stopped")

;; re-show data and summary records, but don't lose the selection
outlinelevel "data"

debug
___ ENDPROCEDURE .print_ogs_batches ____________________________________________

___ PROCEDURE .print_poe_batches _______________________________________________
global collector, order_aggregate
order_aggregate = ""
collector = ""

window waswindow
field «1SpareNumber»
groupup

Field Order
FormulaFill ?(info("summary") =0,«»,collector+assign((sandwich("",collector,¶)+«»)[1,info("summary") =0],"collector"))

;; just show summary records
outlinelevel "1"

Field Con
FormulaFill "Batch " + str(«1SpareNumber») + " Summary"

;; run the summary records through a special macro in the totaller
loop
    order_aggregate = Order
    order_aggregate=replace(order_aggregate, "(","")
    order_aggregate=replace(order_aggregate,")","")
    order_aggregate=replace(order_aggregate," lbs.","")
    openfile "MT Totaller"
    openfile "&@order_aggregate"
    if is_early = "true"
        call .create_early_batch_summ
    else
        call .create_batch_summaries
    endif
    downrecord
until info("stopped")

;; re-show data and summary records, but don't lose the selection
outlinelevel "data"
___ ENDPROCEDURE .print_poe_batches ____________________________________________

___ PROCEDURE .print_this_backorder ____________________________________________
goform "backorder"
if val(BackOrder)>0 
printonerecord
goform "ogspagecheck"

else 
goform "ogspagecheck"
endif
___ ENDPROCEDURE .print_this_backorder _________________________________________

___ PROCEDURE .reexport_oos ____________________________________________________
local outs, allouts, order_len
outs=""

SELECTALL


debug

;; select items in 45orderedogs where the item was out of stocked, 0 filled
window "45orderedogs"
REMOVEALLSUMMARIES
SELECTALL
SELECT comment CONTAINS "out-of-stock"
if info("empty")
else
    ;; hang onto the order numbers that were found, 
    arrayselectedbuild outs, ¬, "", str(OrderNo)
endif


;; go back to the tally
window waswindow

;; select all OGS orders containing "out-of-stock" OR in the list of order numbers just deleted from 45orderedogs OR POE orders containing OGS out of stocks
SELECT (Order CONTAINS "out-of-stock" AND OrderNo < 400000) 
or (OGSOrder CONTAINS "out-of-stock" AND OrderNo > 600000 AND OrderNo < 700000) 
or (arraycontains(outs,str( OrderNo),¬) AND (OrderNo < 400000 or OrderNo > 600000)) 
or (Notes1 contains "cancel" and (OrderNo < 400000 or OrderNo > 600000))
if info("empty")
    window "45orderedogs"
    rtn
endif

;; create a NEW array containing all of the order numbers now selected in tally
;; (may be more than were originally selected in 45orderedogs)
arrayselectedbuild allouts, ¬, "", str(OrderNo)

;; go back to 45orderedogs
window "45orderedogs"

;; select all items on all of the orders in the allouts array
SELECT arraycontains(allouts,str( OrderNo),¬)
if info("empty")
    message "Weird, this shouldn't be empty. Abort!"
    stop
endif

;; delete these orders--all items, not just the out of stocked ones.
SELECTREVERSE
REMOVEUNSELECTED

;; run through those records in tally and put those orders into 45orderedogs

debug

window waswindow
neworder=""
order=""
firstrecord
loop

    order=OGSOrder
    order_len = arraysize(extract(order,¶,1),¬)
    
    case order_len = 11
        ArrayFilter order,neworder,¶,ShipCode+¬+str(OrderNo)+¬+str(Sequence)+¬+import()
    case order_len = 12
        ArrayFilter order,neworder,¶,ShipCode+¬+str(OrderNo)+¬+str(Sequence)+¬+extract(import(),¬,1)+¬+extract(import(),¬,3)+¬+extract(import(),¬,4)+¬+extract(import(),¬,5)+¬+extract(import(),¬,6)+¬+extract(import(),¬,7)+¬+extract(import(),¬,8)+¬+extract(import(),¬,9)+¬+extract(import(),¬,10)+¬+extract(import(),¬,11)+¬+extract(import(),¬,12)
    case Notes1 contains "cancel"
        neworder = ShipCode+¬+str(OrderNo)+¬+str(Sequence)+¬+""+¬+""+¬+""+¬+""+¬+""+¬+"order cancelled"
    endcase

    window "45orderedogs:secret"
    openfile "+@neworder"
    window waswindow
    downrecord
    order=""
    neworder=""
stoploopif OrderNo>700000
until info("stopped")

window "45orderedogs"

Field qty
Formulafill ?(comment contains "out-of-stock",0,qty)

___ ENDPROCEDURE .reexport_oos _________________________________________________

___ PROCEDURE .refund __________________________________________________________
local vhow,vbal,vtransfer,canbatch
canbatch=""
global addpay

addpay=(-1)*«BalDue/Refund»
  if «BalDue/Refund»<.05 and «BalDue/Refund»>0
 goto justapenny
 endif
 
getscrap "How are you issuing this refund? (cc, gc, ch, cash)"

case ((clipboard() contains "cc" or clipboard() contains "cred") and «BalDue/Refund» < 0) 
    ;; rebills are never batched
    canbatch="false"
	   vhow="Credit_Card"
case ((clipboard() contains "cc" or clipboard() contains "cred") and today()-OrderPlaced≥180 and «BalDue/Refund» > 0) 
    ;; old transactions needing a refund
	   Message "Transaction is too old to refund. Issue a check instead."
    stop
case ((clipboard() contains "cc" or clipboard() contains "cred") and (today()-OrderPlaced<180 and (OrderNo >= 500000 and OrderNo < 520000) or (OrderNo >= 300000 and OrderNo<320000) or (OrderNo≥600000 and OrderNo<620000)) and «BalDue/Refund» > 0)
    ;; paper orders that are less than six months old, so the cards can be refunded, but CAN'T be batched
    canbatch="false"
	   vhow="Credit_Card"
case (clipboard() contains "cc" or clipboard() contains "cred")
    ;; any remaining CC transactions are newer than 6 months and are web orders, so they can be batched
    canbatch="true"
	   vhow="Credit_Card"
case (clipboard() contains "gc" or clipboard() contains "gift")
    vhow="Gift_Certificate"
case (clipboard() contains "ca" or clipboard() contains "$")
    vhow="Cash"
case (clipboard() contains "ch" or clipboard() contains "√")
    vhow="Check"
case clipboard() contains "tr"
    vhow="Transfer"
    gettext "To which order are you transferring this payment?",vtransfer
    «Notes2»=«Notes2»+¶+"Refund transferred to payment for order "+vtransfer
endcase

vbal=«BalDue/Refund»
 «BalDue/Refund»=«BalDue/Refund»+addpay
 If «AddPay1»=0
    «AddPay1»= addpay
    «DatePay1»=today()
    «MethodPay1»=vhow
else
    if «AddPay2»=0
        «AddPay2»= addpay
        «DatePay2»=today()
        «MethodPay2»=vhow
    else
        if «AddPay3»=0
            «AddPay3»=addpay
            «DatePay3»=today()
            «MethodPay3»=vhow
        else
            if «AddPay4»=0
            «AddPay4»=addpay
            «DatePay4»=today()
            «MethodPay4»=vhow
            else
                if «AddPay5»=0
                «AddPay5»=addpay
                «DatePay5»=today()
                «MethodPay5»=vhow
                else
                    if «AddPay6»=0
                    «AddPay6»=addpay
                    «DatePay6»=today()
                    «MethodPay6»=vhow
                    else 
                    message "This order is crazy. Please consolidate additional payments."
                    stop
                    endif
                endif
           endif
        endif
    endif
endif

case vhow="Credit_Card"
    if wannarunbatchrefunds="true"
        if canbatch="true"
            global raycc
            waswindow=info("windowname")
            arraylinebuild raycc,¶,"45ogstally",¬+¬+str(OrderNo)+¬+str(Z)+¬+Email+¬+str(GrTotal)+¬+str(abs(addpay))+¬+str(«Auth_Code»)+¬+CreditCard+¬+ExDate
            window "45credit charges"
            openfile "+@raycc"
            save
            window waswindow
        else
            if vbal > 0
	               message "Paper order. Refund manually with Authorize."
            else
                message "Rebill this card manually."
            endif
        endif
    else
        message "Refund/rebill this card manually."
    endif

case vhow="Gift Certificate"
    clipboard()=str(-addpay)
    applescript |||
		      tell application "Firefox"
			     activate
        open location "https://www.fedcoseeds.com/manage_site/create-gift-certificate.php"		  
		      end tell
    |||
    
case vhow="Cash"
    message "Don't forget the refund of $"+str(-addpay)
case vhow="Check"
    window "45Checkregister"
            resynchronize
            lastrecord
            insertbelow
            Name=grabdata("45ogstally", Con)
            «$Amount»=vbal
            Amount=vbal
            Date=today()
            OrderNo=grabdata("45ogstally", OrderNo)
            OpenForm "Check"
            PrintOneRecord dialog   
            CloseWindow    
            save
            WindowToBack "45Checkregister"
endcase

justapenny:
 Paid= Paid+addpay
 «BalDue/Refund»=0
___ ENDPROCEDURE .refund _______________________________________________________

___ PROCEDURE .refundchecksonly ________________________________________________
if «BalDue/Refund»<.05 and «BalDue/Refund»>0
goto justapenny
endif
If «BalDue/Refund»≥3.00
case «AddPay6»>0
message "this order is crazy, please consolidate additional payments"
stop
endcase
            window "45Checkregister"
            resynchronize
            lastrecord
            insertbelow
            Name=grabdata("45ogstally", Con)
            Name=?(grabdata("45ogstally", Con)="", grabdata("45ogstally", Group),Name)
            «$Amount»=grabdata("45ogstally", «BalDue/Refund»)
            Amount=grabdata("45ogstally", «BalDue/Refund»)
            Date=today()
            OrderNo=grabdata("45ogstally", OrderNo)
            OpenForm "Check"
            PrintOneRecord dialog   
            CloseWindow    
            save
            WindowToBack "45Checkregister"
 If «AddPay1»=0
    «AddPay1»= addpay
    «DatePay1»=today()
    «MethodPay1»="Check"
else
    if «AddPay2»=0
        «AddPay2»= addpay
        «DatePay2»=today()
        «MethodPay2»="Check"
    else
        if «AddPay3»=0
            «AddPay3»=addpay
            «DatePay3»=today()
            «MethodPay3»="Check"
        else
            if «AddPay4»=0
            «AddPay4»=addpay
            «DatePay4»=today()
            «MethodPay4»="Check"
            else
                if «AddPay5»=0
                «AddPay5»=addpay
                «DatePay5»=today()
                «MethodPay5»="Check"
                else
                    if «AddPay6»=0
                    «AddPay6»=addpay
                    «DatePay6»=today()
                    «MethodPay6»="Check"
                    else 
                    message "This order is crazy. Please consolidate additional payments."
                    stop
                    endif
                endif
           endif
        endif
    endif
endif
else
message "Don't forget the cash refund of "+str(«BalDue/Refund»)
 If «AddPay1»=0
    «AddPay1»= addpay
    «DatePay1»=today()
    «MethodPay1»="Cash"
else
    if «AddPay2»=0
        «AddPay2»= addpay
        «DatePay2»=today()
        «MethodPay2»="Cash"
    else
        if «AddPay3»=0
            «AddPay3»=addpay
            «DatePay3»=today()
            «MethodPay3»="Cash"
        else
            if «AddPay4»=0
            «AddPay4»=addpay
            «DatePay4»=today()
            «MethodPay4»="Cash"
            else
                if «AddPay5»=0
                «AddPay5»=addpay
                «DatePay5»=today()
                «MethodPay5»="Cash"
                else
                    if «AddPay6»=0
                    «AddPay6»=addpay
                    «DatePay6»=today()
                    «MethodPay6»="Cash"
                    else 
                    message "This order is crazy. Please consolidate additional payments."
                    stop
                    endif
                endif
           endif
        endif
    endif
endif
endif
justapenny:
Paid=Paid-«BalDue/Refund»
«BalDue/Refund»=0
___ ENDPROCEDURE .refundchecksonly _____________________________________________

___ PROCEDURE .RerunPicksheet __________________________________________________
waswindow= info("WindowName")
form= info("FormName")

sub=?(Sub1="Y" and (Sub2="Y" OR Sub2=""),"Y",?(Sub1="Y" and Sub2="N","O",?(Sub1="N" and (Sub2="Y" OR Sub2=""),"E","N")))

openfile "BulbsTotaller"
oldwindow=info("windowname")
window waswindow
order=""
linenu=1
       
loop ;this sets the order variable up to load into the Totaller fields correctly when the picksheet is being run a second time
    order=order+extract(extract(Order,¶,linenu),¬,1)+¬+extract(extract(Order,¶,linenu),¬,2)[1,4]+¬+extract(extract(Order,¶,linenu),¬,3)+¬+
    strip(extract(extract(Order,¶,linenu),¬,4))+¬+
    strip(extract(extract(Order,¶,linenu),¬,6))+¬+extract(extract(Order,¶,linenu),¬,7)+¬+extract(extract(Order,¶,linenu),¬,8)+¬+
    extract(extract(Order,¶,linenu),¬,9)+¬+
    extract(extract(Order,¶,linenu),¬,5)+¶
    linenu=linenu+1
until  extract(extract(Order,¶,linenu),¬,2)=""
           
window "BulbsTotaller"
       
       
openfile "&@order"
call ".rerunpicksheet"
if Order contains "backorder" OR Order contains "ships later"
    if brange=""
        call ".backorder"
    endif
endif
___ ENDPROCEDURE .RerunPicksheet _______________________________________________

___ PROCEDURE .retotal _________________________________________________________
TaxTotal=?(Taxable="Y", TaxTotal,0)
if «Bulk» contains "b" or «Bulk» contains "staff"
    VolDisc=0
else
    VolDisc=float(Subtotal)*float(Discount)
endif

MemDisc=?(MemDisc>0,max(float(Subtotal)*float(.01),.01),0)
AdjTotal=Subtotal-VolDisc-MemDisc
if OrderNo > 500000 and OrderNo < 600000
    ;;BulbsAdjTotal=BulbsSubtotal-(VolDisc+MemDisc)*divzero(BulbsSubtotal,Subtotal)
    BulbsAdjTotal=BulbsSubtotal-float(BulbsSubtotal)*float(Discount)*?(MemDisc>0,float(0.01),1)
endif

GrTotal=0

debug

if GroupMarker contains "P"
    TaxedAmount=taxable
    ;; skips all the rest of the retotaling for group parts
    rtn
endif

if Subtotal > 0
    call ".shipping"
endif

call ".salestax"

if Subtotal=0 and ShipCode≠"D"
    if OrderNo >= 300000 and OrderNo < 400000
        YesNo "Refund Shipping?"
        if clipboard() contains "y"
            «$Shipping»=0
        else
            «$Shipping»=«$Shipping»
        endif
    else
        «$Shipping»=0
    endif
endif

OrderTotal=AdjTotal+«$Shipping»+SalesTax+Surcharge
GrTotal=OrderTotal+Donation+CatalogDefrayment

;; zero out the Actual_Donated_Refund
Paid = Paid + Actual_Donated_Refund
Actual_Donated_Refund = 0

«BalDue/Refund»=Paid-GrTotal

case «BalDue/Refund»=0 and FillDate=today()
    Actual_Donated_Refund=0
case «BalDue/Refund» > 0 and «1stRefund» > 0
        if «Donated_Refund»≥«BalDue/Refund»
            «Actual_Donated_Refund»= «BalDue/Refund» 
            Message "MOFGA Refund is "+pattern(Actual_Donated_Refund,"$#.##")
            Paid=Paid-«BalDue/Refund»
            «BalDue/Refund»=0
        else
            Paid=Paid-«Donated_Refund»
            Actual_Donated_Refund = Donated_Refund
            «BalDue/Refund»=Paid-GrTotal
            Message "MOFGA Refund is "+pattern(Actual_Donated_Refund,"$#.##")
        endif
endcase

if «BalDue/Refund»>0 and info("trigger")="Button.Complete"
    if Subtotal=0 and «$Shipping»>0
        call "nextorder/1"
    else
        call ".refund"
    endif
endif

if «BalDue/Refund»<0 and info("trigger")="Button.Complete"
    call ".baldue"
endif

RealTax=SalesTax

if OrderNo=int(OrderNo)
    Patronage=GrTotal-Donation-CatalogDefrayment-SalesTax
else
    Patronage=0
endif

if groupconsolidation <> "Y" and adjusting_order = 0
    call "nextorder/1"
    groupconsolidation=""
else
    adjusting_order = 0
    groupconsolidation=""
endif


___ ENDPROCEDURE .retotal ______________________________________________________

___ PROCEDURE .salestax ________________________________________________________
debug

case arraycontains(taxstates+","+noship,TaxState,",")=0
    TaxedAmount=0
endcase

;; GROUP ORDERS
case (GroupMarker contains "P" or GroupMarker contains "G")
    if is_completing_group = "true"
        ;; skip this so the tax doesn't get screwed up (already calculated by Cmd-5)
        is_completing_group = ""
    else
        if arraycontains(noship,TaxState,",")=-1 or (TaxState contains "AK" and Notes4 notcontains "AKShippingIsTaxable")
            ;; no tax charged on shipping
            TaxedAmount=TaxedAmount*(1-Discount-?(MemDisc>0,0.01,0))
        else 
            ;; tax charged on shipping
            if AdjTotal = -0.01
                AdjTotal = 0
            endif
            if OrderNo > 500000 and OrderNo < 600000
                ;; bulbs changes - FIXME (can't use bulbstaxtotal for groups)
                TaxedAmount=TaxedAmount*(1-Discount-?(MemDisc>0,0.01,0))+«$Shipping»*divzero(BulbsTaxTotal,BulbsSubtotal)
            else
                TaxedAmount=TaxedAmount*(1-Discount-?(MemDisc>0,0.01,0))+«$Shipping»*divzero(TaxedAmount*(1-Discount-?(MemDisc>0,0.01,0)),AdjTotal)
            endif
        endif
    
        TaxedAmount=?(Taxable="Y",TaxedAmount,0)
    endif
;; NOT GROUP ORDERS
defaultcase
    if arraycontains(noship,TaxState,",")=-1 or (TaxState contains "AK" and Notes4 notcontains "AKShippingIsTaxable")
        ;; no tax charged on shipping
        TaxedAmount=taxable
    else 
        ;; tax charged on shipping
            if OrderNo > 500000 and OrderNo < 600000
                TaxedAmount=taxable+«$Shipping»*divzero(BulbsTaxTotal,BulbsSubtotal)
            else
                TaxedAmount=taxable+divzero(«$Shipping»*taxable,AdjTotal+Surcharge)
            endif
    endif
    
    TaxedAmount=?(Taxable="Y",TaxedAmount,0)
endcase

if OrderNo = int(OrderNo)
    SalesTax=round(float(TaxedAmount)*float(TaxRate)+.0001,.01)
    StateTax=round(float(TaxedAmount)*float(StateRate)+.0001,.01)
    CountyTax=round(float(TaxedAmount)*float(CountyRate)+.0001,.01)
    CityTax=round(float(TaxedAmount)*float(CityRate)+.0001,.01)
    SpecialTax=round(float(TaxedAmount)*float(SpecialRate)+.0001,.01)
endif


;;; figure out what vgroupvermonttax is used for in Bulbs tally
___ ENDPROCEDURE .salestax _____________________________________________________

___ PROCEDURE .selectPOrdersCompToday __________________________________________
select OrderNo < 400000 and OrderNo = int(OrderNo) and ShipCode contains "P" 
and Status contains "c" and PickedUp = "" and FillDate = today() and Email <> ""
if info("empty")
    message "no pickup orders completed today"
endif
___ ENDPROCEDURE .selectPOrdersCompToday _______________________________________

___ PROCEDURE .select_poe_pickup_orders ________________________________________
select OrderNo > 600000 and ShipCode contains "P" and PickedUp notcontains "Yes" and OrderNo = int(OrderNo)

field OrderNo
sortup

openform "POEpickups"
print dialog

field Con
sortup

openform "POEpickups"
print dialog
CloseWindow
___ ENDPROCEDURE .select_poe_pickup_orders _____________________________________

___ PROCEDURE .setcheckers _____________________________________________________
if info("Files") notcontains "Staff"
    openfile "Staff"
    window "Hide This Window"
    window waswindow
endif
zchecker=""
GetScrapOK "Who's the checker?"

If clipboard()=""
YesNo "Is there a checker?"
    if clipboard()="No"
Checker=""
stop
    endif
    If clipboard()="Yes"
call .checkers   
    endif 
endif

if clipboard()≠""
    zcode=clipboard()
    zchecker=lookup("Staff","code",zcode,"checker_name","",0)
        if zchecker≠""
    Checker=zchecker
        endif    
endif

if zchecker=""
Message "You haven't assigned a valid checker, try again!"
call .checkers
endif


;;;;;; maybe??? this last line?

vchecker = Checker
___ ENDPROCEDURE .setcheckers __________________________________________________

___ PROCEDURE .setpullers ______________________________________________________
if info("Files") notcontains "Staff"
    openfile "Staff"
    window "Hide This Window"
    window waswindow
endif
zpuller=""


;; instead of GetScrapOK, do a dropdown menu like the repack file? Can be built programmatically?
;; should filter based on "Active" and "OGS Puller" field

GetScrapOK "Who's the puller?"

If clipboard()=""
    YesNo "Is there a puller?"
    if clipboard()="No"
        Puller=""
        stop
    endif
    If clipboard()="Yes"
        call .pullers   
    endif 
endif

if clipboard()≠""
    zcode=clipboard()
    zpuller=lookup("Staff","Name",zcode,"Name","",0)
        if zpuller≠""
    Puller=zpuller
        endif    
endif

if zpuller=""
Message "You haven't assigned a valid puller, try again!"
call .pullers
endif
___ ENDPROCEDURE .setpullers ___________________________________________________

___ PROCEDURE .shipcodeh _______________________________________________________
local shipmate
OrderComments=OrderComments+¶+"--Original Ship Code "+ShipCode
ShipCode="H"
gettext "Why are we holding this order?",shipmate
«Notes2»=«Notes2»+" Holding for "+shipmate
___ ENDPROCEDURE .shipcodeh ____________________________________________________

___ PROCEDURE .shipping ________________________________________________________
waswindow=info("windowname")
global Shtext, «$Sh», VZ, Vship,v$ship, Vorg, depotprice, cartship
local littleshipcode, prevship
local NumberOfMediumBoxes, NumberOfLargeBoxes, leftoverB


«$Sh»=0
v$ship=«$Shipping»
prevship = «$Shipping»
Vorg= «8SpareNumber»
VZ=Z
littleshipcode=""
cartship=0

If ShipCode Contains "H"
    case OrderComments contains "original shipcode: U"
        ShipCode = "U"
    case OrderComments contains "original shipcode: X"
        ShipCode = "X"
    case OrderComments contains "original shipcode: C"
        ShipCode = "C"
    case OrderComments contains "original shipcode: J"
        ShipCode = "J"
    case OrderComments contains "original shipcode: P"
        ShipCode = "P"
    case OrderComments contains "original shipcode: T"
        ShipCode = "T"
    defaultcase
        if Depot contains "NOFA"
            «$Shipping»=0
        else
            gettext "what is the shipcode?",littleshipcode
            ShipCode = upper(littleshipcode)
        endif
    endcase
Endif

If OrderNo >= 500000 and OrderNo < 600000
    if Pool≠2
        «$Shipping»=?(ShipCode="U" OR ShipCode="L" OR ShipCode="X" OR ShipCode="H",?(BulbsAdjTotal>100,round(BulbsAdjTotal*.12+.0001,.01),12),0)
        «$Shipping»=?(St contains "AK" or St contains "HI" or St contains "AE" or St contains "PR",?(BulbsAdjTotal>100,round(BulbsAdjTotal*.16+.0001,.01),16),«$Shipping»)
    endif
    if (AdjTotal=0 or BulbsAdjTotal=0)
        «$Shipping»=0
    endif
    Field Subtotal
    rtn
endif

Vship=ShippingWt

debug

;; if we've already set a custom shipping charge, don't zero it out
If (ShipCode Contains "U" or ShipCode Contains "X" or ShipCode Contains "G") and ShipCode notcontains "pick"
    if canQuoteFlatShipping="true"
        ;; calculate number of large and medium boxes

        NumberOfMediumBoxes = 0
        NumberOfLargeBoxes = countC

        if countB > 2*countC
            ;; leftoverB is the number of B sized bags left over that can't fit in the boxes with the C's
            leftoverB = countB - 2*countC;
            if (leftoverB = 6)
                NumberOfMediumBoxes = 2;
            else
                if ( (leftoverB mod 5) = 0 or (leftoverB mod 5) = 4 )
                    NumberOfLargeBoxes = NumberOfLargeBoxes + int((leftoverB + 1)/5);
                else
                    NumberOfLargeBoxes = NumberOfLargeBoxes + int(leftoverB/5);
                    NumberOfMediumBoxes = 1;
                endif
            endif
        endif
        
        «$Sh» = NumberOfMediumBoxes*mediumFlatBoxCost + NumberOfLargeBoxes*largeFlatBoxCost
        
    else
        Openfile "45shiplookup"
        Find VZ≥ZipBegin And VZ≤ZipEnd
        Case Zone≥10
            if prevship = 0
                «$Sh»=0
            else
                «$Sh»=prevship
                message "This shipping charge was set manually. If the order changed, the shipping may need to change too."
            endif
        Case Vship=0
    	       «$Sh»=0
        Case Vship≤ 2
    	       «$Sh»=«≤2»
        Case Vship ≤ 5
    	       «$Sh»=«≤5»
        Case Vship≤15
    	       «$Sh»=«≤15»
        Case Vship≤25
    	       «$Sh»=«≤25»
        Case Vship≤35
    	       «$Sh»=«≤35»
        Case Vship≤45
    	       «$Sh»=«≤45»
        Case Vship<200 
 		         «$Sh»=Vship*«>45»
        Case Vship>=200
            if prevship = 0
                «$Sh»=0
            else
                «$Sh»=prevship
                message "This shipping charge was set manually. If the order changed, the shipping may need to change too."
            endif
        EndCase
    endif
    
    window waswindow
    
    ;; Don't want to show this popup when printing picksheets
    if «$Sh»=0 and ShippingWt>0 and PickSheet <> ""
        if ShipCode notcontains "C"
            needquotes = needquotes + str(OrderNo) + ¶
        else
            gettext "How Much is Shipping?",Shtext
            «$Sh»=val(Shtext)
        endif
    endif
    
    «$Shipping»=«$Sh»
    Field Subtotal
Else
    If ShipCode Contains "C" And  «$Shipping»=0 and OrderNo=int(OrderNo)
        Field «$Shipping»
        EditCell
        Field Subtotal
    Else
        If ShipCode Contains "J"
        
        debug
            
            global dcode
            dcode=""
            dcode = Depot
            if dcode contains "ME" and ShippingWt>880
                if ShippingWt≤2000
                «$Sh»=88.00
                else
                Field «$Shipping»
                editcell
                field Subtotal
                endif
            endif
            if (dcode contains "MA" or dcode contains "NH" or dcode contains "RI") and ShippingWt>850
                if ShippingWt≤2000
                «$Sh»=102.00
                else
                Field «$Shipping»
                editcell
                field Subtotal
                endif
            endif
            if (dcode contains "VT" or dcode contains "CT") and ShippingWt>870
                if ShippingWt≤2000
                «$Sh»=112.00
                else
                Field «$Shipping»
                editcell
                field Subtotal
                endif
            endif
            if (dcode contains "NY") and ShippingWt>830
                if ShippingWt≤2000
                «$Sh»=133.00
                else
                Field «$Shipping»
                editcell
                field Subtotal
                endif
            endif
            if (dcode contains "NJ") and ShippingWt>830
                if ShippingWt≤2000
                «$Sh»=175.00
                else
                Field «$Shipping»
                editcell
                field Subtotal
                endif
            endif
            if (dcode contains "PA") and ShippingWt>1060
                if ShippingWt≤2000
                «$Sh»=175.00
                else
                Field «$Shipping»
                editcell
                field Subtotal
                endif
            endif
            if dcode <> "" and dcode notcontains "go ask Alice!"
                Openfile "45shipdepotlookup"
                Find depot_code = dcode
                if info("found")
                    depotprice=«price_per_pound»
                    window waswindow
                    «$Sh»=Vship*depotprice
                else
                    window waswindow
                    message "Depot code not found. Check for typos or invisible characters, and try again."
                endif   
            else
               Depot="go ask Alice!"
            endif      
            
            if «$Sh» < 3
                «$Sh» = 3
            endif
    
            window waswindow
            «$Shipping»=«$Sh»
            Field Subtotal            
            
        Else «$Shipping» =«$Shipping»
            Field Subtotal
        Endif
    Endif
EndIf

___ ENDPROCEDURE .shipping _____________________________________________________

___ PROCEDURE .staffpricing ____________________________________________________
global order, waswindow, truck
truck=""
waswindow=info("windowname")
Discount=0
case info("formname")="ogspagecheck"
    window "45ogscomments.linked:secret"
    Synchronize
    window "45ogscomments:secret"
    openfile "&&" + "45ogscomments.linked"
    openfile "NewOGSTotaller"
    window waswindow
    order=Order
    window "NewOGSTotaller:secret"
    openfile "&@order"
    call ".staffpricing"
case info("formname")="mtpagecheck"
    if extract(extract(Order,¶,1),¬,3) contains "1" or 
        extract(extract(Order,¶,1),¬,3) contains "2" or
        extract(extract(Order,¶,1),¬,3) contains "3" or
        extract(extract(Order,¶,1),¬,3) contains "4"
        message "Run picksheetprint first"
        stop
    endif
    openfile "MT Totaller"
    window waswindow
    order=Order
    order=replace(order, " lbs.","")
    order=replace(order, "(","")
    order=replace(order,")","")
    vship=ShippingWt
    f=Bulk
    truck=ShipCode
    window "MT Totaller:secret"
    openfile "&@order"

    call ".staffpricing"
endcase

___ ENDPROCEDURE .staffpricing _________________________________________________

___ PROCEDURE .todayspickups ___________________________________________________
;select the orders
select ShipCode="P" and PickUpDate=date(|||today|||) and OrderNo<400000 and OrderNo = int(OrderNo)
;open the form
openform "todayspickups"
;print
print dialog
;close the form
closewindow

___ ENDPROCEDURE .todayspickups ________________________________________________

___ PROCEDURE .updateshipping __________________________________________________
waswindow=info("windowname")

case
ShipCode="C" and «UPSTracking#»=""
openfile "45truckinginvoices"
addrecord
OrderNo=str(grabdata("45ogstally",OrderNo))
Carrier=grabdata("45ogstally",TruckingCo)
PRO=str(grabdata("45ogstally",«UPSTracking#»))
DateShipped=today()
St=grabdata("45ogstally",Sta)
ChargedCust=grabdata("45ogstally",«$Shipping»)
local expectations,accessorize
expectations=""
accessorize=""
gettext "Expected Charge?",expectations
if expectations≠""
ExpectedCharge=val(expectations)
endif
gettext "Any Accessorials?",accessorize
Accessorials=accessorize
save

case
ShipCode="C" and «UPSTracking#»≠""
local ono
ono=OrderNo
openfile "45shipping"
selectall
find «OrderNo»=ono
if info("found")=0
lastrecord
insertbelow
Contact=grabdata("45ogstally", Con)
Group=grabdata("45ogstally", Group)
«Sh Add1»=grabdata("45ogstally", SAd)
City=grabdata("45ogstally", Cit)
St=grabdata("45ogstally", Sta)
Zip=grabdata("45ogstally", Z)
«OrderNo»=grabdata("45ogstally", OrderNo)
«UPS#»=grabdata("45ogstally", «UPSTracking#»)+grabdata("45ogstally", «TruckingCo»)
«C#»=str(grabdata("45ogstally", «C#»))
Date=today()
Weight=grabdata("45ogstally", ShippingWt)
endif
openfile "45truckinginvoices"
selectall
find OrderNo=str(ono)
if info("found")=0
addrecord
OrderNo=str(grabdata("45ogstally",OrderNo))
Carrier=grabdata("45ogstally",TruckingCo)
PRO=str(grabdata("45ogstally",«UPSTracking#»))
DateShipped=today()
St=grabdata("45ogstally",Sta)
ChargedCust=grabdata("45ogstally",«$Shipping»)
local expectations,accessorize
expectations=""
accessorize=""
gettext "Expected Charge?",expectations
if expectations≠""
ExpectedCharge=val(expectations)
endif
gettext "Any Accessorials?",accessorize
Accessorials=accessorize
save
endif

defaultcase
openfile "45shipping"
lastrecord
insertbelow
Contact=grabdata("45ogstally", Con)
Group=grabdata("45ogstally", Group)
«Sh Add1»=grabdata("45ogstally", SAd)
City=grabdata("45ogstally", Cit)
St=grabdata("45ogstally", Sta)
Zip=grabdata("45ogstally", Z)
«OrderNo»=grabdata("45ogstally", OrderNo)
«UPS#»=grabdata("45ogstally", «UPSTracking#»)+grabdata("45ogstally", «TruckingCo»)
«C#»=str(grabdata("45ogstally", «C#»))
Date=today()
Weight=grabdata("45ogstally", ShippingWt)
endcase
window waswindow

window waswindow


___ ENDPROCEDURE .updateshipping _______________________________________________

___ PROCEDURE .view_group_coversheet ___________________________________________
if OrderNo > 500000 and OrderNo < 600000
    openform "multicover.bulbs"
else
    openform "multicover.ogs"
endif
___ ENDPROCEDURE .view_group_coversheet ________________________________________

___ PROCEDURE .windowtoback ____________________________________________________
selectall
windowtoback "45ogstally:facilitator view"
___ ENDPROCEDURE .windowtoback _________________________________________________

___ PROCEDURE .zippick _________________________________________________________
local deadline
;getscrap "what is last sequence number?"
;deadline=val(clipboard())
;select «Sequence»>0 and «Sequence»<deadline
getscrap "beginning zipcode"
selectwithin Z≥val(clipboard())
getscrap "ending zipcode"
selectwithin Z≤val(clipboard())
selectwithin ShipCode="U" or ShipCode="X"

___ ENDPROCEDURE .zippick ______________________________________________________

___ PROCEDURE sync & sortup/0 __________________________________________________
ono=OrderNo
ReSynchronize
field OrderNo
sortup
find OrderNo=ono

;find OrderNo≥39000

;if info("found")
;    uprecord
;else
;    lastrecord
;endif

___ ENDPROCEDURE sync & sortup/0 _______________________________________________

___ PROCEDURE nextorder/1 ______________________________________________________
local staff,pullers,checkers
staff=""
pullers=""
checkers=""



if str(OrderNo) contains "."
    downrecord
else
        getscrap "Next!"
        SelectAll
        Numb=val(clipboard())
        if info("formname")="ogspagecheck"
            Numb=?(Numb<10000, Numb+320000,Numb)
            find OrderNo=Numb
        endif
        If info("formname")="mtpagecheck"
            Numb=?(Numb<100000, Numb+600000, Numb)
            find OrderNo=Numb
        endif
        If info("formname")="bulbspagecheck"
            Synchronize
            Numb=?(Numb<100000, Numb+500000, Numb)
            find OrderNo=Numb
        endif
        if Status="" or ((OrderNo > 500000 and OrderNo < 600000) and Status <> "Com")
            if OrderNo < 500000 or OrderNo > 600000
            ;; OGS and POE orders
                if vpuller=""
                    message "set puller"
                endif
                if vchecker=""
                    message "set checker"
                endif
                Puller=vpuller
                Checker=vchecker
                «Puller2»=""
                «Checker2»=""               
            else
            ;; Bulbs orders
                if vpuller=""
                    message "set pullers"
                endif
                if vchecker=""
                    message "set checker"
                endif
                if Puller=""
                    Puller=vpuller
                else
                    if «Puller2»=""
                        «Puller2»=vpuller
                    endif
                endif
                if Checker=""
                    Checker=vchecker
                else
                    if «Checker2»=""
                        «Checker2»=vchecker
                    endif
                endif
            endif  
        endif
        define cname, ""
        define cgroup, ""
        define findall, ""
        superobject "orderfinder", "FillList"
        superobject "OrderList", "FillList", ""
        drawobjects
        
endif

stop
___ ENDPROCEDURE nextorder/1 ___________________________________________________

___ PROCEDURE setcheckerauto ___________________________________________________
if info("formname")="bulbspagecheck"
    case info("trigger")="puller1popup"
        vpuller=Puller
    case info("trigger")="puller2popup"
        vpuller=«Puller2»
    case info("trigger")="checker1popup"
        vchecker=Checker
    case info("trigger")="checker2popup"
        vchecker=«Checker2»
    endcase
else
    vpuller=Puller
    vpullertwo=""
    vchecker=Checker
    vcheckertwo=""
endif



___ ENDPROCEDURE setcheckerauto ________________________________________________

___ PROCEDURE justlocate/4 _____________________________________________________
fileglobal orderno
orderno=""

If info("selected") < info("records")
SelectAll
EndIf

case info("formname")="ogsinput"
    gettext "Next!", orderno
    Numb=val(orderno)
    Numb=?(Numb<10000, Numb+320000,Numb)
    find OrderNo=Numb
    stop
case info("formname")="ogspagecheck"
    getscrap "Next!"
    Numb=val(clipboard())
    Numb=?(Numb<10000, Numb+320000,Numb)
    find OrderNo=Numb
case info("formname")="mtpagecheck"
    gettext "Next!", orderno
    Numb=val(orderno)
    Numb=?(Numb<100000, Numb+600000, Numb)
    find OrderNo=Numb
case info("formname")="bulbspagecheck"
    getscrap "Next!"
    Numb=val(clipboard())
    Numb=?(Numb<100000, Numb+500000,Numb)
    find OrderNo=Numb
defaultcase
    gettext "Next!", orderno
    Numb=val(orderno)
    find OrderNo=Numb
endcase

case Numb >= 600000
    goform "mtpagecheck"
case Numb >= 300000 and Numb < 400000
    goform "ogspagecheck"
case Numb >= 500000 and Numb < 600000
    goform "bulbspagecheck"
endcase


case info("formname")="enter mail"
    Numb=?(val(ono)<10000, val(ono)+30000, val(ono))
    find OrderNo=Numb
    ono=""
    stop
endcase

define cname, ""
define cgroup, ""
define findall, ""
superobject "orderfinder", "FillList"
superobject "OrderList", "FillList", ""
drawobjects


___ ENDPROCEDURE justlocate/4 __________________________________________________

___ PROCEDURE multipagetotal/5 _________________________________________________
global groupconsolidation

groupconsolidation="Y"
rayg="M"

NoUndo
getscrap "Order to consolidate"
Numb=val(clipboard())

case info("formname")="ogspagecheck"
    Numb=?(Numb<100000, Numb+300000,Numb)
case info("formname")="mtpagecheck"
    Numb=?(Numb<100000, Numb+600000,Numb)
case info("formname")="bulbspagecheck"
    Numb=?(Numb<100000, Numb+500000,Numb)
endcase

Select int(OrderNo) = Numb
if info("selected")=1 or info("selected") = info("records")
    message "This is not a group order."
    stop
endif

Field OrderNo
SortUp

FirstRecord
Subtotal=0
BulbsSubtotal=0
BulbsTaxTotal=0

Order=""
BackOrder=""
ShippingWt=0
TaxedAmount=0

Field Subtotal
Total

Field BulbsSubtotal
Total

Field BulbsTaxTotal
Total

Field TaxedAmount
Total

Field ShippingWt
Total

Field «5SpareMoney»
Total

Field CatalogDefrayment
Total

gr=""

firstrecord
if (OrderNo > 500000 and OrderNo < 600000)
    loop 
        if OrderNo <> int(OrderNo)
            gr=gr+?(Status="B/O", str(OrderNo)[".",-1]+" backorder"+¶, ?(Status="Com", str(OrderNo)[".",-1]+" com"+¶, str(OrderNo)[".",-1]+" -"+¶))
        endif
        downrecord
    until info("stopped")
else
    loop 
        gr=?(Status="B/O", str(OrderNo)[".",-1]+" "+BackOrder+¶+gr, gr)
        downrecord
    until info("stopped")
endif

gr=?(gr="" or (gr notcontains "backorder" and gr notcontains "-"), rep(chr(32),10)+"Complete",gr)

local bulbssub, bulbstaxt

LastRecord
sub=Subtotal
bulbssub=BulbsSubtotal
bulbstaxt=BulbsTaxTotal
vdefray=CatalogDefrayment
grouptax=TaxedAmount
Vship=ShippingWt
disctotal=«5SpareMoney»

RemoveSummaries 7

FirstRecord
Order=gr
Subtotal=sub
BulbsSubtotal=bulbssub
BulbsTaxTotal=bulbstaxt
CatalogDefrayment=vdefray
MemDisc=?(MemDisc>0,.01*Subtotal,0)
TaxedAmount=grouptax
ShippingWt=Vship
«5SpareMoney»=disctotal
disctotal=0
Checker=vchecker
Puller=vpuller
if OrderNo≥500000 and OrderNo<600000
«Checker2»=vcheckertwo
«Puller2»=vpullertwo
endif

call .retotal



___ ENDPROCEDURE multipagetotal/5 ______________________________________________

___ PROCEDURE moose_puller_stats/n _____________________________________________
Select Status contains "c" and «OrderNo» >= 600000 and Order notcontains "complete"

Field "Puller"
GroupUp

Field "FillDate"
GroupUp by Day

Field "TaxTotal"
Count

Field "ShippingWt"
Total

Field "PullError"
Total

Field "Puller"
Propagate

CollapseToLevel "1"

export "moose_stats.txt", Puller+¬+datestr(FillDate)+¬+str(TaxTotal)+¬+str(ShippingWt)+¬+str(«PullError»)+¶

RemoveSummaries 7
___ ENDPROCEDURE moose_puller_stats/n __________________________________________

___ PROCEDURE multiorderreprint/6 ______________________________________________
global groupprint
groupprint="Y"

firstrecord
local groupdisc,groupwt,groupship, groupsub, grouptaxrate, grouptaxedamount, sumtaxedamount, groupsalestax, grouptaxstate
local partadjustedtotal, partshipping, parttotal, parttax
tray=""
oray=""
gr=""
groupdisc=?(MemDisc=0,Discount,Discount+.01)
groupwt=ShippingWt
groupship=«$Shipping»
groupsub=Subtotal
grouptaxrate=TaxRate
groupsalestax=SalesTax
grouptaxedamount=TaxedAmount
grouptaxstate=TaxState

field TaxedAmount
Total
lastrecord
sumtaxedamount = TaxedAmount - grouptaxedamount
removesummaries 7

firstrecord
downrecord

debug

case arraycontains(taxstates,grouptaxstate,",")
    ;; this state taxes shipping
    if OrderNo > 500000 and OrderNo < 600000
        loop
            partadjustedtotal = Subtotal*(1-groupdisc)
            parttax = divzero(TaxedAmount,sumtaxedamount)*groupsalestax
            partshipping = divzero(Subtotal,groupsub)*groupship
            parttotal = partadjustedtotal + parttax + partshipping
            gr=Con[1,20]
            tray=
            rep(chr(32),2)+pattern(OrderNo,"#.###")[".",-1]
            +rep(chr(32),20-length(gr))
            +gr
            +rep(chr(32),10-length(pattern(Subtotal, "$#,.##")))
            +pattern(Subtotal, "$#,.##")
            +rep(chr(32),10-length(pattern(partadjustedtotal, "$#,.##")))
            +pattern(partadjustedtotal, "$#,.##")
            +rep(chr(32),9-length(pattern(parttax,"$#,.##")))
            +pattern(parttax,"$#,.##")
            +rep(chr(32),9-length(pattern(partshipping,"$#,.##")))
            +pattern(partshipping,"$#,.##")
            +rep(chr(32),12-length(pattern(parttotal,"$#,.##")))
            +pattern(parttotal,"$#,.##")
            +¶
            oray=oray+tray
            downrecord
        until info("stopped")
    else
        loop
            partadjustedtotal = Subtotal*(1-groupdisc)
            parttax = divzero(TaxedAmount,sumtaxedamount)*groupsalestax
            partshipping = divzero(Subtotal,groupsub)*groupship
            gr=Con[1,20]
            tray=
            rep(chr(32),2)+pattern(OrderNo,"#.###")[".",-1]
            +rep(chr(32),20-length(gr))
            +gr
            +rep(chr(32),10-length(pattern(Subtotal, "$#,.##")))
            +pattern(Subtotal, "$#,.##")
            +rep(chr(32),10-length(pattern(partadjustedtotal, "$#,.##")))
            +pattern(partadjustedtotal, "$#,.##")
            +rep(chr(32),9-length(pattern(parttax,"$#,.##")))
            +pattern(parttax,"$#,.##")
            +rep(chr(32),9-length(pattern(divzero(ShippingWt,groupwt)*groupship,"$#,.##")))
            +pattern(divzero(ShippingWt,groupwt)*groupship,"$#,.##")
            +rep(chr(32),6-length(str(ShippingWt)))
            +str(ShippingWt)+" lbs."+¶
            oray=oray+tray
            downrecord
        until info("stopped")
    endif

case arraycontains(noship,grouptaxstate,",")
    message "this state does not tax shipping"
    if OrderNo > 500000 and OrderNo < 600000
        loop
            partadjustedtotal = Subtotal*(1-groupdisc)
            parttax = divzero(TaxedAmount,sumtaxedamount)*groupsalestax
            partshipping = divzero(Subtotal,groupsub)*groupship
            parttotal = partadjustedtotal + parttax + partshipping
            gr=Con[1,20]
            tray=
            rep(chr(32),2)+pattern(OrderNo,"#.###")[".",-1]
            +rep(chr(32),20-length(gr))
            +gr
            +rep(chr(32),10-length(pattern(Subtotal, "$#,.##")))
            +pattern(Subtotal, "$#,.##")
            +rep(chr(32),10-length(pattern(partadjustedtotal, "$#,.##")))
            +pattern(partadjustedtotal, "$#,.##")
            +rep(chr(32),9-length(pattern(parttax,"$#,.##")))
            +pattern(parttax,"$#,.##")
            +rep(chr(32),9-length(pattern(partshipping,"$#,.##")))
            +pattern(partshipping,"$#,.##")
            +rep(chr(32),12-length(pattern(parttotal,"$#,.##")))
            +pattern(parttotal,"$#,.##")
            +¶
            oray=oray+tray
            downrecord
        until info("stopped")
    else
        loop
            partadjustedtotal = Subtotal*(1-groupdisc)
            parttax = divzero(TaxedAmount,sumtaxedamount)*groupsalestax
            partshipping = divzero(ShippingWt,groupwt)*groupship ;; NOTE this calculation is different for states that tax shipping
            gr=Con[1,20]
            tray=
            rep(chr(32),2)+pattern(OrderNo,"#.###")[".",-1]
            +rep(chr(32),20-length(gr))
            +gr
            +rep(chr(32),10-length(pattern(Subtotal, "$#,.##")))
            +pattern(Subtotal, "$#,.##")
            +rep(chr(32),10-length(pattern(partadjustedtotal, "$#,.##")))
            +pattern(partadjustedtotal, "$#,.##")
            +rep(chr(32),9-length(pattern(parttax,"$#,.##")))
            +pattern(parttax,"$#,.##")
            +rep(chr(32),9-length(pattern(partshipping,"$#,.##")))
            +pattern(partshipping,"$#,.##")
            +rep(chr(32),6-length(str(ShippingWt)))
            +str(ShippingWt)+" lbs."+¶
            oray=oray+tray
            downrecord
        until info("stopped")
    endif

defaultcase
    message "this state doesn't tax shit"
    if OrderNo > 500000 and OrderNo < 600000
        loop
            partadjustedtotal = Subtotal*(1-groupdisc)
            parttax = 0
            partshipping = divzero(Subtotal,groupsub)*groupship
            parttotal = partadjustedtotal + parttax + partshipping
            gr=Con[1,20]
            tray=
            rep(chr(32),2)+pattern(OrderNo,"#.###")[".",-1]
            +rep(chr(32),20-length(gr))
            +gr
            +rep(chr(32),10-length(pattern(Subtotal, "$#,.##")))
            +pattern(Subtotal, "$#,.##")
            +rep(chr(32),10-length(pattern(partadjustedtotal, "$#,.##")))
            +pattern(partadjustedtotal, "$#,.##")            
            +rep(chr(32),4)
            +"$0.00"
            +rep(chr(32),9-length(pattern(partshipping,"$#,.##")))
            +pattern(partshipping,"$#,.##")
            +rep(chr(32),12-length(pattern(parttotal,"$#,.##")))
            +pattern(parttotal,"$#,.##")
            +¶
            oray=oray+tray
            downrecord
        until info("stopped")    
    else
        loop
            partadjustedtotal = Subtotal*(1-groupdisc)
            parttax = 0
            partshipping = divzero(ShippingWt,groupwt)*groupship ;; NOTE this calculation is different for states that tax shipping
            gr=Con[1,20]
            tray=
            rep(chr(32),2)+pattern(OrderNo,"#.###")[".",-1]
            +rep(chr(32),20-length(gr))
            +gr
            +rep(chr(32),10-length(pattern(Subtotal, "$#,.##")))
            +pattern(Subtotal, "$#,.##")
            +rep(chr(32),10-length(pattern(partadjustedtotal, "$#,.##")))
            +pattern(partadjustedtotal, "$#,.##")            
            +rep(chr(32),4)
            +"$0.00"
            +rep(chr(32),9-length(pattern(partshipping,"$#,.##")))
            +pattern(partshipping,"$#,.##")
            +rep(chr(32),6-length(str(ShippingWt)))
            +str(ShippingWt)+" lbs"+¶
            oray=oray+tray
            downrecord
        until info("stopped")
    endif

endcase

firstrecord
«6SpareText» = oray

if OrderNo >= 500000 and OrderNo < 600000
    openform "multicover.bulbs"
else
    openform "multicover.ogs"
endif
debug
printonerecord dialog
;Print 
closewindow
ono=int(OrderNo)
;selectall
gr=""
find OrderNo=ono

___ ENDPROCEDURE multiorderreprint/6 ___________________________________________

___ PROCEDURE justmoose/7 ______________________________________________________
select OrderNo≥600000

___ ENDPROCEDURE justmoose/7 ___________________________________________________

___ PROCEDURE Update Tally Numbers/9 ___________________________________________
global order, neworder
local order_len
local allwindows, numwindows, n, onewindowname

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; This new and improved macro captures the sold AND filled numbers
;; from orders in the tally. It looks at all of the orders each time it's run
;; and rebuilds 45orderedogs from scratch. The macro is able to look
;; for OGS items on POE orders (not Bulbs orders, yet).
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;; go to the datasheet. If the datasheet is not currently open, 
;; go to the active window and switch it to the datasheet
gosheet

;; create a list of all of the currently open windows in the tally
allwindows = listwindows("45ogstally")

;; Since we just ran "gosheet," the datasheet will be the active
;; window, and will be the first one in the list of windows. We're
;; going to close all of the windows in this list, so before we start
;; that, we want to delete the datasheet from the list.

if arraysize(allwindows,¶) > 1
    allwindows = arraydelete(allwindows,1,1,¶)
    numwindows = arraysize(allwindows,¶)

    ;; loop through the list of windows (except for the datasheet)
    ;; and close them all
    n = 0
    loop
        onewindowname = array(allwindows,n+1,¶)
        window onewindowname
        closewindow
        n = n+1
        numwindows = numwindows - 1
        stoploopif numwindows = 0
    while forever

endif

waswindow=info("windowname")
 
;;==========================================

select  OrderNo < 400000 and OrderNo > 300000 and ShipCode notcontains "D" and OGSOrder contains ')'
noyes "check POE orders?"
PayAttentionToPoe = clipboard()

if PayAttentionToPoe = "Yes"
    selectadditional OrderNo >= 600000 and OGSOrder <> ''
 endif
 
noyes "check Bulbs Orders?"
PayAttentionToBulbs = clipboard()

if PayAttentionToBulbs = "Yes"
    selectadditional OrderNo < 600000 and OrderNo > 500000 and OGSOrder <> ''
 endif

field OrderNo
sortup

openfile "45orderedogs"
DeleteAll

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;; Since orders are often filled out of sequence, currently the best way to capture
;; all filled orders and changes to past orders is to rerun the macro for ALL of 
;; the orders every time we want to update the inventory.
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

window waswindow
neworder=""
order=""
firstrecord

debug

NoShow
loop
    order=OGSOrder
    order_len = arraysize(extract(order,¶,1),¬)
    
    case order_len = 11
        ArrayFilter order,neworder,¶,ShipCode+¬+str(OrderNo)+¬+str(Sequence)+¬+import()
    case order_len = 12
        ArrayFilter order,neworder,¶,ShipCode+¬+str(OrderNo)+¬+str(Sequence)+¬+extract(import(),¬,1)+¬+extract(import(),¬,3)+¬+extract(import(),¬,4)+¬+extract(import(),¬,5)+¬+extract(import(),¬,6)+¬+extract(import(),¬,7)+¬+extract(import(),¬,8)+¬+extract(import(),¬,9)+¬+extract(import(),¬,10)+¬+extract(import(),¬,11)+¬+extract(import(),¬,12)
    defaultcase
        message "something weird in the OGSOrder field for " + str(OrderNo) + "!" + " Array length is " + str(order_len) + "."
        stop
    endcase

    window "45orderedogs:secret"
    openfile "+@neworder"
    window waswindow
    downrecord
    order=""
    neworder=""
until info("stopped")
EndNoShow

;;========================================================

window waswindow
selectwithin Status = ''

window "45orderedogs"

Field qty
Formulafill ?(comment contains "out-of-stock",0,qty)

Select lookupselected("45ogstally","OrderNo",OrderNo,"OrderNo",0,0)
Field fill
Fill 0

selectall

Field "IDNumber"
GroupUp
Field "qty"
Total
Field "fill"
Total
CollapseToLevel "1"

Save
Openfile "45ogscomments.linked"
gosheet
message "Finished NEW Tally Export Macro"

___ ENDPROCEDURE Update Tally Numbers/9 ________________________________________

___ PROCEDURE additionalpay/å __________________________________________________
local addpay, vhow,vorderref,vorderpay

if automaticwriteoff = "true"
    addpay=-«BalDue/Refund»
    vhow="Write_Off"
    goto skipbuttons
endif

debug
 
case info("trigger") = "Button.Undo AddPay1"
    Paid=Paid-AddPay1
    «BalDue/Refund»=«BalDue/Refund»-AddPay1
    AddPay1=0
    DatePay1=0
    MethodPay1=""
    stop 
case info("trigger") = "Button.Undo AddPay2"
    Paid=Paid-AddPay2
    «BalDue/Refund»=«BalDue/Refund»-AddPay2
    AddPay2=0
    DatePay2=0
    MethodPay2=""
    stop 
case info("trigger") = "Button.Undo AddPay3"
    Paid=Paid-AddPay3
    «BalDue/Refund»=«BalDue/Refund»-AddPay3
    AddPay3=0
    DatePay3=0
    MethodPay3=""
    stop 
case info("trigger") = "Button.Undo AddPay4"
    Paid=Paid-AddPay4
    «BalDue/Refund»=«BalDue/Refund»-AddPay4
    AddPay4=0
    DatePay4=0
    MethodPay4=""
    stop 
case info("trigger") = "Button.Undo AddPay5"
    Paid=Paid-AddPay5
    «BalDue/Refund»=«BalDue/Refund»-AddPay5
    AddPay5=0
    DatePay5=0
    MethodPay5=""
    stop 
case info("trigger") = "Button.Undo AddPay6"
    Paid=Paid-AddPay6
    «BalDue/Refund»=«BalDue/Refund»-AddPay6
    AddPay6=0
    DatePay6=0
    MethodPay6=""
    stop 
 endcase

 GetScrap  "What's the additional payment?"
 addpay=val(clipboard())
 getscrap "How was it paid? (cc, gc, ch, cash, wr)"
 ;; avoid issuing CC refunds for POE and Bulbs paper orders because they do batch refunds. No restrictions for OGS, unless they start to do batch refunds.
;; case ((clipboard() contains "cc" or clipboard() contains "cred") and (today()-OrderPlaced≥90 or (OrderNo >= 500000 and OrderNo < 520000) or (OrderNo >= 300000 and OrderNo<320000) or (OrderNo≥600000 and OrderNo<620000)))
;; case ((clipboard() contains "cc" or clipboard() contains "cred") and (today()-OrderPlaced≥90 or (OrderNo >= 500000 and OrderNo < 520000) or (OrderNo≥600000 and OrderNo<620000)))
case (clipboard() contains "cc" or clipboard() contains "cred")
	   vhow="Credit_Card"
 case (clipboard() contains "gc" or clipboard() contains "gift")
    vhow="Gift_Certificate"
 case (clipboard() contains "ch" or clipboard() contains "√")
    vhow="Check"
 case (clipboard() contains "cash" or clipboard() contains "$")
    vhow="Cash"
 case (clipboard() contains "wr")
    vhow="Write_Off"
 endcase
 
 case clipboard() contains "tr"
    vhow="Transfer"
    vorderref=str(OrderNo)
    gettext "To which order are you transferring this payment?",vorderpay
    «Notes2»=«Notes2»+¶+"Payment transferred to order "+vorderpay
 endcase
 
 skipbuttons:
 
 Paid= Paid+addpay
 
 «BalDue/Refund»=«BalDue/Refund»+addpay
 If «AddPay1»=0
    «AddPay1»= addpay
    «DatePay1»=today()
    «MethodPay1»=vhow
else
    if «AddPay2»=0
        «AddPay2»= addpay
        «DatePay2»=today()
        «MethodPay2»=vhow
    else
        if «AddPay3»=0
            «AddPay3»=addpay
            «DatePay3»=today()
            «MethodPay3»=vhow
        else
            if «AddPay4»=0
            «AddPay4»=addpay
            «DatePay4»=today()
            «MethodPay4»=vhow
            else
                if «AddPay5»=0
                «AddPay5»=addpay
                «DatePay5»=today()
                «MethodPay5»=vhow
                else
                    if «AddPay6»=0
                    «AddPay6»=addpay
                    «DatePay6»=today()
                    «MethodPay6»=vhow
                    else 
                    message "This order is crazy. Please consolidate additional payments."
                    stop
                    endif
                endif
           endif
        endif
    endif
endif

if vhow="Credit_Card"
    if addpay < 0
        message "Remember to refund this card manually."
    else
        message "Remember to rebill this card manually."
    endif
endif

if vhow="Transfer"
find OrderNo=val(vorderpay)
    if info("found")
    addpay=-addpay
    «Notes2»=«Notes2»+¶+"Payment transferred from order "+str(vorderref)
     Paid= Paid+addpay
 
 «BalDue/Refund»=«BalDue/Refund»+addpay
 If «AddPay1»=0
    «AddPay1»= addpay
    «DatePay1»=today()
    «MethodPay1»=vhow
else
    if «AddPay2»=0
        «AddPay2»= addpay
        «DatePay2»=today()
        «MethodPay2»=vhow
    else
        if «AddPay3»=0
            «AddPay3»=addpay
            «DatePay3»=today()
            «MethodPay3»=vhow
        else
            if «AddPay4»=0
            «AddPay4»=addpay
            «DatePay4»=today()
            «MethodPay4»=vhow
            else
                if «AddPay5»=0
                «AddPay5»=addpay
                «DatePay5»=today()
                «MethodPay5»=vhow
                else
                    if «AddPay6»=0
                    «AddPay6»=addpay
                    «DatePay6»=today()
                    «MethodPay6»=vhow
                    else 
                    message "This order is crazy. Please consolidate additional payments."
                    stop
                    endif
                endif
           endif
        endif
    endif
endif
yesno
"Does this look right?"
if clipboard() contains "n"
stop
endif
find OrderNo=val(vorderref)
else
message "Can't find the order you're using for this transfer"
endif
endif
    

___ ENDPROCEDURE additionalpay/å _______________________________________________

___ PROCEDURE PicksheetPrint/π _________________________________________________
global oldwindow, oldfile, pickno, taxable
taxable=0
needquotes=""

oldfile=info("DatabaseName") 
oldwindow=info("windowname")


case info("formname")="ogspagecheck"
    if OrderNo>600000
        goform "mtpagecheck"
        call ".mtpicksheets"
    endif
    
    waswindow=info("windowname")
    call ".ogspicksheet"
    
    
case info("formname")="mtpagecheck"
    if OrderNo<600000
        yesno "Print MT Tubers?"
        if clipboard()="Yes"
            find OrderNo=600001
        else
            stop
        endif
    endif
    
    waswindow=info("windowname")
    call ".mtpicksheets"
    
case info("formname")="bulbspagecheck"
    if OrderNo > 600000 or OrderNo < 500000
        yesno "Print Bulbs?"
        if clipboard()="Yes"
            find OrderNo=500001
        else
            stop
        endif
    endif
    
    if info("trigger")="Button.run this picksheet"
        just_this_one="Yes"
    else
        just_this_one="No"
    endif
    
    waswindow=info("windowname")
    call ".bulbspicksheets"
    
endcase

___ ENDPROCEDURE PicksheetPrint/π ______________________________________________

___ PROCEDURE POE Early Pcksht Prnt/e __________________________________________
global oldwindow, oldfile, pickno
needquotes=""

oldfile=info("DatabaseName") 
oldwindow=info("windowname")

case info("formname")="ogspagecheck"
    if OrderNo>600000
        goform "mtpagecheck"
        call ".mtearlypicksheets"
    endif
    
    message "this macro is just for Moose Tubers early shipment items"
        
case info("formname")="mtpagecheck"
    if OrderNo<600000
        yesno "Print MT Tubers?"
        if clipboard()="Yes"
        find OrderNo=600001
        else
        stop
        endif
    endif
    
    waswindow=info("windowname")
    call ".mtearlypicksheets"
endcase

___ ENDPROCEDURE POE Early Pcksht Prnt/e _______________________________________

___ PROCEDURE ReprintPicksheet/® _______________________________________________
global order, waswindow,oldwindow

waswindow=info("windowname")

case info("formname")="ogspagecheck"
    if OrderNo>600000 or OrderNo < 300000
        message "You are on the wrong page"
        stop
    endif
    call ".ogsreprint"

case info("formname")="mtpagecheck"
    if OrderNo<600000
        message "You are on the wrong page"
        stop
    endif
    call ".mtreprint"
    
case info("formname")="bulbspagecheck"
    if OrderNo<500000 or OrderNo >= 600000
        message "You are on the wrong page"
        stop
    endif
    call ".bulbsreprint"
endcase
___ ENDPROCEDURE ReprintPicksheet/® ____________________________________________

___ PROCEDURE BatchReprintPicksheets ___________________________________________
YesNo "Do you really want to reprint the picksheets for all "+str(info("Selected"))+" orders that are selected?"
if clipboard()="Yes"
firstrecord
call "ReprintPicksheet/®"
    loop
    downrecord
    call "ReprintPicksheet/®"
    until  info("EOF") 
endif
___ ENDPROCEDURE BatchReprintPicksheets ________________________________________

___ PROCEDURE preprint_addon ___________________________________________________
local prenum, presize, precount
Message "This allows the addition of items to an order before the picksheet has been run. The math will be refigured when the picksheet is run."

if PickSheet≠""
    message "This only works before the picksheet has been run."
    stop
endif

loop
    prenum=""
    presize=""
    precount=0
    GetScrap "What's the item?"
    prenum=clipboard()
    GetScrap "What size?"
    presize=upper(clipboard())
    GetScrap "How many"
    precount=clipboard()
    
    if val(prenum)=0 or val(presize)>0 or val(precount)=0
        message "You need to fill in something for each of the item, size and count parameters. Please try again."
        stop
    endif
    
    message prenum+presize+"-"+precount
    
    if prenum≠"" and presize≠"" and precount≠""
        Order=Order+¶+¬+prenum+¬+presize+¬+precount+¬+"preadd"
    endif
    
until prenum=""

___ ENDPROCEDURE preprint_addon ________________________________________________

___ PROCEDURE change/h _________________________________________________________
global order, waswindow, oldtotal, oldship, truck, orig_order, orig_order_stripped, next_item, vedad, whodis
local istruck

istruck=«8SpareText»

orig_order=""
orig_order_stripped=""
next_item=1
oldtotal=AdjTotal
oldship=ShippingWt
vedad=TaxState
waswindow=info("windowname")

if info("trigger") = "Button.AdjustOrder" or info("trigger") = "Button.Adjust Order" or whodis contains "NOFA_settle_up"
    adjusting_order = 1
    «1SpareText» = datepattern(today(),"mm/dd/yy") + " " + timepattern(now(),"HH:MM AM/PM")
else
    adjusting_order = 0
endif

case info("formname")="ogspagecheck"
    if OrderNo≥600000 or OrderNo < 300000
        stop
    endif

    if info("trigger")="Button.Bulk"
        Bulk = "bulk"
    endif
   
    window "45ogscomments.linked"
    ReSynchronize
    window "45ogscomments"
    openfile "&&" + "45ogscomments.linked"
    openfile "NewOGSTotaller"
    window waswindow
    order=Order
    orig_order = «Original Order»
    Sh=ShipCode
    loop
        orig_order_stripped = orig_order_stripped + array(array(orig_order,next_item,¶),2,¬) +¬+ array(array(orig_order,next_item,¶),4,¬) + ¶
        next_item = next_item + 1
    until arraysize(orig_order, ¶)
    window "NewOGSTotaller"
    openfile "&@order"
    ;stop
        
    if info("trigger")="Button.Bulk"
        call ".nofapricing" 
    else
        call ".adjustment"
    endif
    
    if info("trigger")="Button.AdditionalOrder"
        SpareComments="addon"
        call ".additionalorder"
    endif
    
    order = ""

case info("formname")="mtpagecheck"
    if info("trigger")="Button.Bulk"
        Bulk="bulk"
    endif
    
    if OrderNo<600000
        stop
    endif
    
    openfile "MT Totaller"
    window waswindow
    
    if extract(extract(Order,¶,1),¬,3) contains "1" or 
        extract(extract(Order,¶,1),¬,3) contains "2" or
        extract(extract(Order,¶,1),¬,3) contains "3" or
        extract(extract(Order,¶,1),¬,3) contains "4" or
        extract(extract(Order,¶,1),¬,3) contains "5" or
        extract(extract(Order,¶,1),¬,3) contains "6" or
        extract(extract(Order,¶,1),¬,3) contains "7" or
        extract(extract(Order,¶,1),¬,3) contains "8" or
        extract(extract(Order,¶,1),¬,3) contains "9"
        YesNo "Do you want to adjust this order?"
        if clipboard()="Yes"
            f=Bulk
            order=Order
            order=replace(order, " lbs.","")
            order=replace(order, "(","")
            order=replace(order,")","")
            order=replace(order,"–",¬) ;;//This replaces the n-dash which is what the office uses to connect Item with size.
            order=replace(order,"-",¬) ;;//This replaces a regular dash should you happen to use that.
            arrayfilter order, order, ¶, arrayinsert(extract(order,¶,seq()),4,1,¬)
            window "MT Totaller"
            openfile "&@order"
            ;stop
            call ".newadjustment"
        endif
        ;;message "Run picksheetprint first"
        order = ""
        stop
    endif
        
    window waswindow
    order=Order
    order=replace(order, "(","")
    order=replace(order,")","")
    order=replace(order," lbs.","")
    vship=ShippingWt
    f=Bulk
    truck=ShipCode
    istruck =val(«8SpareText»)
    window "MT Totaller"
    openfile "&@order"
    
    debug
    ;stop
        if istruck = 1
            call ".truckadjustment"
        else
            call ".adjustment"
        endif
        
    order = ""

case info("formname")="bulbspagecheck"
    if OrderNo<500000 or OrderNo > 600000
        stop
    endif
    
    if PickSheet=""
        ;; 2022 going to try allowing orders to be adjusted before printing
        YesNo "Order hasn't been printed. Adjust it anyway?"
        if clipboard() = "Yes"
            global vnum, vedad, sub
            vedad = Sta
            vnum = round(OrderNo,.001)
            order=Order
            «Original Order»=Order ;; save a copy of the original order
        
            sub=?(Sub1="Y" and (Sub2="Y" OR Sub2=""),"Y",?(Sub1="Y" and Sub2="N","O",?(Sub1="N" and (Sub2="Y" OR Sub2=""),"E","N")))
            ono=OrderNo
            openfile "45ogstaxable"
            window "Hide This Window"
            openfile "BulbsTotaller"
            openfile "&@order"
            call ".preprint-adjust"
            stop
        else
            stop
        endif
    endif
    
    if OriginalPicksheet=""
        OriginalPicksheet=PickSheet
    endif
    
    sub=?(Sub1="Y" and (Sub2="Y" OR Sub2=""),"Y",?(Sub1="Y" and Sub2="N","O",?(Sub1="N" and (Sub2="Y" OR Sub2=""),"E","N")))
    
    
    openfile "BulbsTotaller"
    openfile "45ogstaxable"
    window "Hide This Window"

    window waswindow
    order=""
    linenu=1
    loop
        order=order+extract(extract(Order,¶,linenu),¬,1)+¬+extract(extract(Order,¶,linenu),¬,2)+¬+extract(extract(Order,¶,linenu),¬,3)+¬+
            extract(extract(Order,¶,linenu),¬,4)+¬+strip(extract(extract(Order,¶,linenu),¬,6))+¬+
            extract(extract(Order,¶,linenu),¬,7)+¬+extract(extract(Order,¶,linenu),¬,8)+¬+extract(extract(Order,¶,linenu),¬,9)+¬+
            extract(extract(Order,¶,linenu),¬,5)+¶
            linenu=linenu+1
    until  extract(extract(Order,¶,linenu),¬,2)=""   
    window "BulbsTotaller"
    openfile "&@order"
    call ".recalculate"
    
    if Order contains "backorder" OR Order contains "ships later"
        if brange=""
            call ".backorder"
        endif
    endif
    
    if addon≠""
        message "Please remember to press BACKORDER or COMPLETE to fully update this order's record when you're done."
    endif
    
    addon=""
endcase


whodis = ""
___ ENDPROCEDURE change/h ______________________________________________________

___ PROCEDURE reworkchanges/® __________________________________________________
;; adding this menu item even though it duplicates change/h, because the keyboard shortcut is familiar to Bulbs/Seeds paper workers (Cmd-opt-R)

call "change/h"
___ ENDPROCEDURE reworkchanges/® _______________________________________________

___ PROCEDURE runpaperwork/ç ___________________________________________________
;; adding this menu item even though it duplicates .done, 
;; because the keyboard shortcut is familiar to Bulbs/Seeds paper workers (Cmd-opt-C)
;; and they prefer having a keyboard shortcut (in addition to button click option)
;; for accessibility reasons

call .done

___ ENDPROCEDURE runpaperwork/ç ________________________________________________

___ PROCEDURE Export Onion Orders ______________________________________________
local newarray, waswindow

waswindow = info("windowname")

Select «2SpareDate» = date("") and  «OrderNo» >= 600000 and ( «Order» contains "7490" or «Order» contains "7500" or «Order» contains "7510" or «Order» contains "7519" or «Order» contains "7520" or «Order» contains "7550")
if info("empty")
    message "no new onion orders"
    stop
endif


arrayselectedbuild newarray,¶,"",str(OrderNo)+¬+ShipCode+¬+str(«C#»)+¬+Con+¬+Group+¬+Email+¬+MAd+¬+City+¬+St+¬+str(Zip)+¬+replace(CustomerComments,¶,'---') +¬+replace(OrderComments,¶,'---')

openfile "onion export helper"
openfile "&@newarray"

 call ".bring_in_onions"

window waswindow

debug

Field «2SpareDate»
FormulaFill today()

message "Done! Look for the 'onion orders.tsv' file; fill in group part addresses and check for customer comments!"

___ ENDPROCEDURE Export Onion Orders ___________________________________________

___ PROCEDURE Export Sw Potato Orders __________________________________________
local newarray, waswindow

waswindow = info("windowname")

Select «3SpareDate» = date("") and  «OrderNo» >= 600000 and ( «Order» contains "7997" or «Order» contains "7998" or «Order» contains "7999")
;; Select «OrderNo» >= 600000 and ( «Order» contains "7997" or «Order» contains "7998" or «Order» contains "7999")
if info("empty")
    message "no new sweet potato orders"
    stop
endif

arrayselectedbuild newarray,¶,"",str(OrderNo)+¬+ShipCode+¬+str(«C#»)+¬+Con+¬+Group+¬+Email+¬+MAd+¬+City+¬+St+¬+str(Zip)+¬+Telephone+¬+replace(CustomerComments,¶,"---") +¬+replace(OrderComments,¶,"---")

openfile "sweet potato export helper"
openfile "&@newarray"

call ".bring_in_sw_potatoes_new"

debug

window waswindow

Field «3SpareDate»
FormulaFill today()

window "sweet potato export helper"

message "Fill in group part addresses and check for customer comments in this file and then save it!"

___ ENDPROCEDURE Export Sw Potato Orders _______________________________________

___ PROCEDURE mailexport _______________________________________________________


local importarray
synchronize
selectall

;select OrderNo>60000 and MoosePicksheetPrinted contains "P"
;and ShipCode notcontains "T" 
;selectwithin Order contains "7990" or Order contains "7995" 
;or Order contains "7955" or Order contains "7970" or Order contains "7972" or Order contains "7975" or Order contains "7980" or Order contains "7990" or Order contains "7995" or Order contains "7997" or Order contains "7999"
;stop
;select OrderNo=int(OrderNo) and OrderNo<60000 and Status notcontains "Com" 
;and ShippingWt<15


;selectadditional Status contains "Com" and FillDate>datevalue(2016,12,27)
select OrderNo>100000 ;and Status contains "Com" and FillDate>today()-30
selectadditional (OrderNo > 500000 and OrderNo < 600000)



;selectadditional OrderNo<60000 and Status notcontains "" and FillDate>today()-30
;selectadditional OrderPlaced>today()-15
;selectadditional Status contains "B/O"
;selectadditional «c/s»>0

ono=OrderNo
importarray=""

arrayselectedbuild importarray,chr(10),"",?(Group≠"",Group+","+Con," "+","+Con)+","+MAd+","+City+","+St+","+pattern(Zip,"#####")+","+Email+","+str(OrderNo)+Con[1,1]+City[1,1]+chr(13)

importarray="Group"+","+"Con"+","+"MAd"+","+"City"+","+"St"+","+"Zip"+","+"Email"+","+"OrderNo"+Con[1,1]+City[1,1]+chr(13)+chr(10)+importarray

select OrderNo=ono
export "mailaddress.csv", importarray
selectall
___ ENDPROCEDURE mailexport ____________________________________________________

___ PROCEDURE (details) ________________________________________________________

___ ENDPROCEDURE (details) _____________________________________________________

___ PROCEDURE AdjustMofgaRefund ________________________________________________
getscrap "What should Paid be reset to before you reComplete the order?"
Paid=val(clipboard())
«BalDue/Refund»=Paid-GrTotal
getscrap "How much of their refund should go to MOFGA?"
Actual_Donated_Refund=val(clipboard())
«BalDue/Refund»=«BalDue/Refund»-Actual_Donated_Refund

___ ENDPROCEDURE AdjustMofgaRefund _____________________________________________

___ PROCEDURE ChangePaid _______________________________________________________
getscrap "New Paid"
Paid=val(clipboard())
«BalDue/Refund»=Paid-GrTotal
if Paid≠«1stPayment»+AddPay1+AddPay2+AddPay3+AddPay4+AddPay5+AddPay6
    message "The amount you entered doesn't match the payment(s) recorded. Please be sure this is correct. Talk to someone if you're unsure, it can avoid bookkeeping headaches. Thanks. Alice"
endif
___ ENDPROCEDURE ChangePaid ____________________________________________________

___ PROCEDURE ChangeShipCode ___________________________________________________
local newship
newship=""

case OrderNo >= 500000 and OrderNo < 600000
    gettext "What's the new ship code (U,P,X,H,D or L)?" newship
    if newship≠""
        ShipCode=upper(newship)
    endif
    if ShipCode≠"U" and ShipCode≠"P" and ShipCode≠"X" and ShipCode≠"H" and ShipCode≠"D" and ShipCode≠"L"
        message ShipCode+" is not a standard code. Please correct if necessary."
    endif
case (OrderNo >= 300000 and OrderNo < 400000) or (OrderNo >= 600000 and OrderNo < 700000)
    gettext "What's the new ship code (U,P,T,X,H,C, J, Q, or D)?" newship
    if newship≠""
        ShipCode=upper(newship)
    endif
    if ShipCode≠"U" and ShipCode≠"P" and ShipCode≠"X" and ShipCode≠"H" and ShipCode≠"D" and ShipCode≠"C" and ShipCode≠"J" and ShipCode≠"Q"
        message ShipCode+" is not a standard code. Please correct if necessary."
    endif
    if ShipCode="H"
    message "Please leave a note specifying the reason for the hold"
    endif
    if ShipCode="Q" and DeliveryDate≤today()
    message "Please use Schedule Order button to correct scheduled date or change ship code to reflect actual method of shipment"
    endif
    if ShipCode="P" and PickUpDate≤today() and Status=""
    message "Please enter anticipated pickup date if known"
    endif
endcase
___ ENDPROCEDURE ChangeShipCode ________________________________________________

___ PROCEDURE ChangeDiscount ___________________________________________________
getscrap "what's the new discount?"
Discount=val(clipboard())
call ".retotal"
___ ENDPROCEDURE ChangeDiscount ________________________________________________

___ PROCEDURE MooseSalesToDate _________________________________________________
;; this used to be the moosetodate macro in the orders file.
;; starting in FY 45 I'm intending to run this macro out of the 
;; tally after appending new batches of orders to the tally. -SAO 8/3

local newarray, exported
newarray=""
exported=""
selectwithin Order <> ""
firstrecord
loop

arrayfilter Order, newarray, ¶, arraydelete(extract(Order,¶,seq()),1,1,¬)

newarray=replace(newarray, "–",¬)

arrayfilter newarray, newarray, ¶, arraydelete(extract(newarray, ¶, seq()), 4,2,¬)
arrayfilter newarray, newarray, ¶, arraydelete(extract(newarray, ¶, seq()), 7,1,¬)
arrayfilter newarray, newarray,¶, ShipCode+¬+str(OrderNo)+¬+str(Sequence)+¬+Sub1+¬+Sub2
    +¬+datepattern(OrderPlaced,"mm/dd/yyyy")+¬+datepattern(EntryDate,"mm/dd/yyyy")+¬+import()+¬+
    arraystrip(stripchar(CustomerComments,"!ÿ "),¬)+¬+arraystrip(stripchar(OrderComments,"!ÿ "),¬)+¬+
    arraystrip(stripchar(Con,"!ÿ "),¬)+¬+Telephone+¬+arraystrip(arraystrip(Email,¶),¬)
exported=exported+newarray+¶
newarray=""
downrecord
until info("stopped")

debug

openfile "mooseboughttodate"
openfile "&@exported"
call "distribution"
___ ENDPROCEDURE MooseSalesToDate ______________________________________________

___ PROCEDURE moosevertical ____________________________________________________
local newarray, exported

newarray=""
exported=""
Synchronize
;; since we no longer backorder every item on ginger and turmeric orders, the "filled" numbers are off when we include B/O statuses.
;;Select OrderNo>600000 and Status≠""
Select OrderNo>600000 and Status contains "c"
firstrecord

NoShow
loop
    arrayfilter Order, newarray, ¶, arraydelete(extract(Order,¶,seq()),1,2,¬)

    arrayfilter Order, newarray, ¶ arraydelete(extract(Order,¶,seq()),5,1,¬)
    newarray= replace(newarray, "–",¬)
    
    arrayfilter newarray, newarray,¶, str(Sequence)+¬+ShipCode+¬+str(OrderNo)+¬+Status+¬+datestr(FillDate)+¬+Sub1+¬+Sub2+¬+import()
    exported=exported+newarray+¶
    newarray=""
    downrecord
until info("stopped")
EndNoShow


openfile "moosevertical"
deleteall
openfile "&@exported"

___ ENDPROCEDURE moosevertical _________________________________________________

___ PROCEDURE ShipLookup _______________________________________________________
Numb=int(OrderNo)
if ShipCode="U" OR ShipCode="X" or ShipCode="C"
    OpenFile "45shipping"
    selectall
    ReSynchronize
    select «OrderNo»=Numb OR «C#»= str(Numb)
else
YesNo "This order is a pickup. Do you still want to check shipping?"
    if clipboard()="Yes"
        OpenFile "45shipping"
        selectall
    ReSynchronize
        select «OrderNo»=Numb OR «C#»= str(Numb)
    endif
endif
if   info("Selected") < info("Records") 
    ShowPage
    stop
else
    message "nothing found"
endif

___ ENDPROCEDURE ShipLookup ____________________________________________________

___ PROCEDURE ---- _____________________________________________________________

___ ENDPROCEDURE ---- __________________________________________________________

___ PROCEDURE backorder_search _________________________________________________
selectall
GetScrapOK "Which backorder item do you want to search for?"
if clipboard()≠""
select Backorder contains clipboard()
    if  info("Selected") =  info("Records") 
    message "nothing found"
    endif
endif

___ ENDPROCEDURE backorder_search ______________________________________________

___ PROCEDURE FindItem/ß _______________________________________________________
selectall
GetScrapOK "what item?"
select Order contains clipboard()
if  info("Selected") =  info("Records") 
    message "nothing found"
endif
___ ENDPROCEDURE FindItem/ß ____________________________________________________

___ PROCEDURE CancelOrder ______________________________________________________
global taxable

zcancel="No"
NoYes "Are you sure you want to cancel this entire order?"
if clipboard()="Yes"
    zcancel="Yes"
    taxable = 0
    Checker=""
    Puller=""
    Order=""
    PickSheet=""
    BackOrder=""
    ShippingWt=0
    TaxTotal=0
    Subtotal=0
    SalesTax=0
    AdjTotal=0
    VolDisc=0
    MemDisc=0
    Surcharge=0
    «$Shipping»=0
    TaxedAmount=0
    OrderTotal=0
    Donation=0
    CatalogDefrayment=0
    Membership=0
    GrTotal=0
    RealTax=0
    Patronage=0
    TaxRate=0
    Notes1=Notes1+" "+"Order cancelled "+datepattern(today(),"mm/dd/yy")+"."
    call .done
endif
zcancel="No" 
___ ENDPROCEDURE CancelOrder ___________________________________________________

___ PROCEDURE (doodah) _________________________________________________________

___ ENDPROCEDURE (doodah) ______________________________________________________

___ PROCEDURE get tracking _____________________________________________________
global track
if info("files") notcontains "45shipping"
openfile "45shipping"
endif
track=""
track=lookupalldouble("45shipping","C#",«C#Text»,"OrderNo","UPS#",¶,¬)
message track
track=""
___ ENDPROCEDURE get tracking __________________________________________________

___ PROCEDURE poe search sh requests ___________________________________________
Select OrderComments contains "hold"
or CustomerComments contains "hold"
or Notes1 contains "hold"
or Notes2 contains "hold"
or Notes3 contains "hold"
or Notes4 contains "hold"
or OrderComments contains "ship"
or CustomerComments contains "ship"
or Notes1 contains "ship"
or Notes2 contains "ship"
or Notes3 contains "ship"
or Notes4 contains "ship"
or OrderComments contains "may"
or CustomerComments contains "may"
or Notes1 contains "may"
or Notes2 contains "may"
or Notes3 contains "may"
or Notes4 contains "may"
or OrderComments contains "june"
or CustomerComments contains "june"
or Notes1 contains "june"
or Notes2 contains "june"
or Notes3 contains "june"
or Notes4 contains "june"
or OrderComments contains "april"
or CustomerComments contains "april"
or Notes1 contains "april"
or Notes2 contains "april"
or Notes3 contains "april"
or Notes4 contains "april"
or OrderComments contains "march"
or CustomerComments contains "march"
or Notes1 contains "march"
or Notes2 contains "march"
or Notes3 contains "march"
or Notes4 contains "march"
or OrderComments contains "late"
or CustomerComments contains "late"
or Notes1 contains "late"
or Notes2 contains "late"
or Notes3 contains "late"
or Notes4 contains "late"
or OrderComments contains "early"
or CustomerComments contains "early"
or Notes1 contains "early"
or Notes2 contains "early"
or Notes3 contains "early"
or Notes4 contains "early"
or OrderComments contains "earliest"
or CustomerComments contains "earliest"
or Notes1 contains "earliest"
or Notes2 contains "earliest"
or Notes3 contains "earliest"
or Notes4 contains "earliest"

Selectwithin OrderNo > 600000 
and ShipCode notcontains "Q" and ShipCode notcontains "C" and ShipCode notcontains "J" and ShipCode notcontains "T" 
and OrderNo = int(OrderNo) ;;and Status notcontains "com"

___ ENDPROCEDURE poe search sh requests ________________________________________

___ PROCEDURE ogs backorder coversheet _________________________________________
global order, order_no, huge_order, searchitem
local num_orders

num_orders = 0
huge_order = ""
searchitem = ""

;; This macro creates a summary report of OGS backorders of a specific item.
;; The report also shows the other backordered items on each order (if there are any)

waswindow = info("windowname")

gettext "Item to search for (xxxx or xxxx-y):",searchitem

Select BackOrder contains searchitem and OrderNo >= 300000 and OrderNo < 400000 and Status like "B/O"
if info("empty")
    message "no backorders for this item(?!?)"
    stop
endif

;; This narrows down selection to exclude group parents, but it keeps group parts
selectwithin OrderNo < 310000 or (OrderNo >= 320000 and OrderNo < 390000) or OrderNo <> int(OrderNo)

;; Opens up the files that will be needed. It is slightly faster to do lookups in
;; an unlinked file, which is why we're opening 45ogscomments and importing
;; the current contents of 45ogscomments.linked.

openfile "NewOGSTotaller"
window "45ogscomments.linked"
Synchronize
openfile "45ogscomments"
openfile "&&" + "45ogscomments.linked"

select Item contains searchitem

window waswindow

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;; Build an array called "huge_order" that contains the contents of the
;; BackOrder field of all of the selected orders, AND the Order number stuck
;; on the end of each line.

Field BackOrder
arrayselectedbuild huge_order, ¶, "", replace(arraystrip(«»,¶),¶,¬+str(OrderNo)+¶)+¬+str(OrderNo)

;; Open the totaller and replace the records with the contents of the huge_order array.

openfile "NewOGSTotaller"
openfile "&@huge_order"
   
call .backordercover

window waswindow

___ ENDPROCEDURE ogs backorder coversheet ______________________________________

___ PROCEDURE ogs generate coversheet __________________________________________
global order, order_no, huge_order
local num_orders

num_orders = 0
huge_order = ""

;; This macro creates a summary coversheet for a selection of orders.
;; It is intended to be used for orders that are already printed (for example,
;; a selection that includes all Tree Sale pickup orders). It has not been tested
;; with unprinted orders.

;; Reminder to make a selection before running this macro. You do not
;; want to run this macro on the whole database :-/

waswindow = info("windowname")
yesno "Have you made your selection of printed orders?"
if clipboard() contains "No"
    "Select the orders before running this macro."
    stop
endif

;; This narrows down selection to exclude group parents, but it keeps group parts
selectwithin OrderNo > 300000 and OrderNo < 400000 and GroupMarker contains "P"

;; Exclude orders that are already completed or backordered. This might need to be
;; changed once I understand what is desired as far as a coversheet for backorders.
selectwithin Status notcontains "c" and Status notcontains "b"

;; Opens up the files that will be needed. It is slightly faster to do lookups in
;; an unlinked file, which is why we're opening 45ogscomments and importing
;; the current contents of 45ogscomments.linked.

openfile "NewOGSTotaller"
window "45ogscomments.linked"
Synchronize
openfile "45ogscomments"
openfile "&&" + "45ogscomments.linked"

window waswindow

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

;; Build an array called "huge_order" that contains the contents of the
;; Order field of all of the selected orders, AND the Order number stuck
;; on the end of each line.

Field Order
arrayselectedbuild huge_order, ¶, "", replace(«»,¶,¬+str(OrderNo)+¶)+¬+str(OrderNo)

;; Run the array through a special macro in NewOGSTotaller, but first get
;; rid of any parentheses or "lbs" that might have been in the Orders.

huge_order=replace(huge_order, "(","")
huge_order=replace(huge_order,")","")
huge_order=replace(huge_order," lbs.","")
openfile "NewOGSTotaller"
openfile "&@huge_order"
        
call .coversheetcollection

window waswindow

___ ENDPROCEDURE ogs generate coversheet _______________________________________

___ PROCEDURE ogsmonthly _______________________________________________________
local seedsselected
select OrderNo = int(OrderNo)
seedsselected=info("selected")
;selectwithin FillDate ≥ month1st((month1st(today())-1)) and FillDate < month1st(today())
; and Status contains "com"
selectwithin Status contains "com"
;if info("selected")=seedsselected
;beep
;stop
;endif
//selectwithin arraycontains(ztaxstates,TaxState,",")


arrayselectedbuild raya,¶,"", "OGS/MT"+¬+str(OrderNo)+¬+Taxable+¬+TaxState+¬+Cit+¬+pattern(Z,"#####")+¬+¬+str(TaxRate)+¬+str(StateRate)+¬+
str(CountyRate)+¬+str(CityRate)+¬+str(SpecialRate)+¬+str(«$Shipping»)+¬+str(AdjTotal)+¬+str(TaxedAmount)+¬+str(SalesTax)+¬+datepattern(FillDate,"Month YYYY")+¬+¬+
str(StateTax)+¬+str(CountyTax)+¬+str(CityTax)+¬+str(SpecialTax)+¬+str(OrderTotal)
    clipboard()=raya
    ;stop
 openfile "WayfairSalesTax"
 openfile "&@raya"  
;call "monthlytotals/t"
___ ENDPROCEDURE ogsmonthly ____________________________________________________

___ PROCEDURE mtreprintfilled __________________________________________________
waswindow=info("windowname")
openform "mtpicksheetfilled"
printonerecord dialog
CloseWindow

if arraysize(rayb,¶)<43
    goto done
else
    openform "mtpicksheet2filled"
    printonerecord dialog
    CloseWindow
endif

done:

window waswindow
___ ENDPROCEDURE mtreprintfilled _______________________________________________

___ PROCEDURE --- ______________________________________________________________

___ ENDPROCEDURE --- ___________________________________________________________

___ PROCEDURE getGroup _________________________________________________________
ono=int(OrderNo)
SelectAll
Select int(OrderNo)=ono

___ ENDPROCEDURE getGroup ______________________________________________________

___ PROCEDURE BackorderCheck ___________________________________________________
selectall
ReSynchronize
GetScrapOK "what item on B/O"
Select Backorder contains clipboard()
___ ENDPROCEDURE BackorderCheck ________________________________________________

___ PROCEDURE Update_Delinked_Comments _________________________________________
openfile "45BulbsComments linked"
ReSynchronize
Field number
sortup
save
closefile
openfile "45BulbsComments"
openfile "&&45BulbsComments linked"
save
MakeSecret
window waswindow

___ ENDPROCEDURE Update_Delinked_Comments ______________________________________

___ PROCEDURE reset to original order __________________________________________
Order=«Original Order»
yesno "Do you want to clear the picksheet?"
if clipboard()="Yes"
    PickSheet=""
endif

yesno "Do you want to clear all checkers and pullers?"
if clipboard()="Yes"
Checker=""
Puller=""
Checker2=""
Puller2=""
endif

yesno "Do you want to clear status?"
if clipboard()="Yes"
Status=""
endif

yesno "Do you want to clear print status?"
if clipboard()="Yes"
print_code=""
in_process=""
endif
___ ENDPROCEDURE reset to original order _______________________________________

___ PROCEDURE Dunum ____________________________________________________________
GetScrapOK "Minimum balance due to search for?"
zdun=0-val(clipboard())
selectall

case info("formname")="ogspagecheck"
    select OrderNo = int(OrderNo) and OrderNo > 300000 and OrderNo < 400000
    selectwithin «BalDue/Refund» ≤ zdun
    selectwithin ShipCode notcontains "D"
    selectwithin Status contains "c"
    selectwithin Notes1 notcontains "barter"
    selectwithin Notes1 notcontains "ignore"
case info("formname")="bulbspagecheck"
    select OrderNo = int(OrderNo) and OrderNo > 500000 and OrderNo < 600000
    selectwithin «BalDue/Refund» ≤ zdun
    selectwithin ShipCode notcontains "D"
    selectwithin Notes1 notcontains "barter"
    selectwithin Notes1 notcontains "ignore"
case info("formname")="mtpagecheck"
    select OrderNo = int(OrderNo) and OrderNo > 600000 and OrderNo < 700000
    selectwithin «BalDue/Refund» ≤ zdun
    selectwithin ShipCode notcontains "D"
    selectwithin Notes1 notcontains "barter"
    selectwithin Notes1 notcontains "ignore"
    selectwithin Status contains "c"
defaultcase
    message "go to one of the pagecheck forms to search within that branch's orders"
    stop
endcase

YesNo "exclude orders marked billed?"
if clipboard()="Yes"
    selectwithin Notes1 notcontains "billed"
endif

___ ENDPROCEDURE Dunum _________________________________________________________

___ PROCEDURE diagnostics ______________________________________________________
global minorderno, maxorderno


if zdiagnose=""
    message "searching for unreworked orders with an o needs work."
    bigmessage "remember to check groups manually (manufactured groups, especially) for weirdnesses (VolDisc, Discount, Paid,1stPayment,1stTotal,1stRefund & BalDue/Refund should all be cleared)."
    zdiagnose="Y"
    message "this macro only syncs the first time in a session"
    Synchronize
    
    ;;
    ;; ask which branch you want to diagnose
    GetScrap "Which branch are you checking? (bulbs, ogs, poe)"
    
    case clipboard() contains "bulbs"
        minorderno = 500000
        maxorderno = 600000
    case clipboard() contains "ogs"
        minorderno = 300000
        maxorderno = 400000
    case clipboard() contains "poe"
        minorderno = 600000
        maxorderno = 700000
    defaultcase
        message "Try that again, but with better spelling."
        zdiagnose=""
        stop
    endcase
    ;;
endif

field OrderNo
sortup



;; needs work
case zlevel=0
    YesNo "1. Select unreworked orders?"
    if clipboard()="Yes"
        select OrderNo > minorderno and OrderNo < maxorderno
        Selectwithin Order contains "b     " and Order notcontains "sub" and 
            Order notcontains "limit" and Order notcontains "pack" and Order notcontains "give"
            and Order notcontains "see RB" and Order notcontains "only"
        Selectadditional Order contains "o     " and Order notcontains "ECO"
        selectwithin Notes1 notcontains "all set"
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=1
            call diagnostics
            stop
        endif
        zlevel=1
        stop
    endif

case zlevel=1
    YesNo "2. Select Order still contains backorder?"
    if clipboard()="Yes"
        select Order contains "backorder" and OrderNo > minorderno and OrderNo < maxorderno and Notes1 notcontains "all set"
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=2
            call diagnostics
            stop
        endif
        zlevel=2
        stop
    endif

case zlevel=2
    YesNo "3. Select Status is blank?"
    if clipboard()="Yes"
        select Status="" and OrderNo > minorderno and OrderNo < maxorderno and Notes1 notcontains "all set"
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=3
            call diagnostics
            stop
        endif
        zlevel=3
        stop
    endif

case zlevel=3
    YesNo "4. Select Status≠Com?"
    if clipboard()="Yes"
        select Status<>"Com" and Notes1 notcontains "all set" and OrderNo > minorderno and OrderNo < maxorderno
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=4
            call diagnostics
            stop
        endif
        zlevel=4
        stop
    endif

case zlevel=4
    YesNo "5. Select Status≠"" AND Paid=0? (typically purchase orders and side deals)"
    if clipboard()="Yes"
        select Status≠"" and Paid=0 and OrderNo = int(OrderNo) and GrTotal≠0 and OrderNo > minorderno and OrderNo < maxorderno and Notes1 notcontains "all set"
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=6
            call diagnostics
            stop
        endif
        zlevel=6
        stop
    endif

;; not sure this is a meaningful test--1stTotal is 0 for many many orders

;;case zlevel=5
;;    YesNo "6. Select «1stTotal»=0 or looks off?"
;;    if clipboard()="Yes"
;;        select «1stTotal»=0 and OrderNo = int(OrderNo)
;;        selectwithin OrderNo > minorderno and OrderNo < maxorderno
;;        selectwithin Notes1 notcontains "all set"
;;        selectwithin Notes1 notcontains "order cancelled"
;;        selectwithin OrderComments notcontains "empty order"
;;    zlevel=6
;;    stop
;;endif

case zlevel=6
    YesNo "6. Select AddPay2≠0?"
    if clipboard()="Yes"
        select AddPay2≠0 and Order notcontains "add" and OrderNo > minorderno and OrderNo < maxorderno and Notes1 notcontains "all set"
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=7
            call diagnostics
            stop
        endif
        zlevel=7
        stop
    endif

case zlevel=7
    YesNo "7. Select BalDue/Refund>0? (we owe them money)"
    if clipboard()="Yes"
        select «BalDue/Refund»>0 and OrderNo > minorderno and OrderNo < maxorderno and Notes1 notcontains "all set" and Status≠""
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=8
            call diagnostics
            stop
        endif
        zlevel=8
        stop
    endif

case zlevel=8
    YesNo "8. Select patron+RealTax≠OrderTotal? (due to donations)"
    if clipboard()="Yes"
        select Status≠"" and Patronage+RealTax≠OrderTotal and OrderNo > minorderno and OrderNo < maxorderno and Notes1 notcontains "all set"
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=9
            call diagnostics
            stop
        endif
        zlevel=9
        stop
    endif

case zlevel=9
    YesNo "9. Select Pd≠1stPayment and Bal<0? (find quirks)"
    if clipboard()="Yes"
        select Status≠"" and Paid≠«1stPayment» and «BalDue/Refund»<0 and «BalDue/Refund»≠-.01 
        and OrderNo > minorderno and OrderNo < maxorderno and Notes1 notcontains "all set"
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=10
            call diagnostics
            stop
        endif
        zlevel=10
        stop
    endif

case zlevel=10
    YesNo "10. Select OrderTotal, Patronage or RealTax ≠ 0 for group pieces?"
    if clipboard()="Yes"
        select ((OrderNo <> int(OrderNo) and OrderTotal <> 0) OR (OrderNo <> int(OrderNo) and Patronage <> 0) OR (OrderNo <> int(OrderNo) and RealTax <> 0)) 
        and OrderNo > minorderno and OrderNo < maxorderno and Notes1 notcontains "all set"
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=11
            call diagnostics
            stop
        endif
        zlevel=11
        stop
    endif

case zlevel=11
    YesNo "11. Select OrderTotal+Donation+CatalogDefrayment+Membership≠GrTotal?" ;does not compute for group pieces as they have no grand total
    if clipboard()="Yes"
        select OrderTotal+Donation+CatalogDefrayment+Membership≠GrTotal and OrderNo = int(OrderNo) 
        and OrderNo > minorderno and OrderNo < maxorderno and Notes1 notcontains "all set"
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=12
            call diagnostics
            stop
        endif
        zlevel=12
        stop
    endif

case zlevel=12
    YesNo "12. Select order with no payments apparently? Typically meaning 1st payment field was not filled in at the office."
    if clipboard()="Yes"
        select «1stPayment»=0 and OrderNo = int(OrderNo) and (AddPay1+AddPay2+AddPay3+AddPay4+AddPay5+AddPay6)<Paid 
        and OrderNo > minorderno and OrderNo < maxorderno and Notes1 notcontains "all set"
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=13
            call diagnostics
            stop
        endif
        zlevel=13
        stop
    endif

case zlevel=13
    YesNo "13. Select RealTax≠SalesTax? (last diagnostic)"; does not compute for group pieces
    if clipboard()="Yes"
        select RealTax≠SalesTax and OrderNo = int(OrderNo) 
        and OrderNo > minorderno and OrderNo < maxorderno and Notes1 notcontains "all set"
        if info("empty")
            message "Looks good! Moving on..."
            zlevel=14
            call diagnostics
            stop
        endif
        zlevel=14
        stop
    endif

case zlevel=14
zlevel=0

endcase
___ ENDPROCEDURE diagnostics ___________________________________________________

___ PROCEDURE single_item_exporter _____________________________________________
;this looks for a specific item and exports that line of data related to that order to the order&line file. Way cool!
local ordernu, itemnu, backline, added
backline=""
added=""
itemnu=""
firstrecord

;this piece gets, configures and selects for the item you're looking for
    GetScrapOK "What's the item number"
    itemnu=clipboard()
    select Order contains itemnu    
   ; stop
   
;this piece extracts the item from the orders
itemnu="*"+itemnu+"*"
linenu=1
ordernu=1
loop
    linenu=arraysearch(Order,itemnu, linenu,¶)
    stoploopif linenu=0
    backline=str(OrderNo)+¬+str(Sequence)+¬+Con+¬+City+¬+St+¬+str(Zip)+¬+Email+¬+Telephone+¬+Sub1+¬+Sub2+¬+extract(Order,¶,linenu)+¶
    added=added+backline
    linenu=linenu+1
until info("empty")

loop
    downrecord
    ordernu=ordernu+1
    linenu=1

    loop
        linenu=arraysearch(Order,itemnu, linenu,¶)
        stoploopif linenu=0
        backline=str(OrderNo)+¬+str(Sequence)+¬+Con+¬+City+¬+St+¬+str(Zip)+¬+Email+¬+Telephone+¬+Sub1+¬+Sub2+¬+extract(Order,¶,linenu)+¶
        added=added+backline
        linenu=linenu+1
    until info("empty")

until  info("EOF") 

;message added+"-yo"

openfile "order&line"
openfile "&@added"
___ ENDPROCEDURE single_item_exporter __________________________________________

___ PROCEDURE unlockrecord _____________________________________________________
forceunlockrecord
___ ENDPROCEDURE unlockrecord __________________________________________________

___ PROCEDURE (testing) ________________________________________________________

___ ENDPROCEDURE (testing) _____________________________________________________

___ PROCEDURE depot ____________________________________________________________
local title
GetScrap "Which Depot?"
title=upper(clipboard())

debug

Select (ShipCode="J" or ShipCode="H") and Depot=title and Status≠"Com"
field OrderNo
sortup
export title+"emails",Email+","
GoSheet

Field ShippingWt 
Total
Field AdjTotal 
Total

copycell
pastecell

Depot=title
GoForm "depotmasterlist"
Message "Save as PDF and Print Hard Copy"
Print Dialog
Print Dialog
GoForm "depotsetcontacts"
Message "Save as PDF"
Print Dialog
gosheet
GetScrap "Who is Trucking?"
«TruckingCo»=clipboard()
Depot=title
FillDate=today()
window "45shipdepotlookup"
if error
    openfile "45shipdepotlookup"
endif
window "45ogstally"
Group=lookup("45shipdepotlookup","depot_code",«Depot»,"Farm","DEPOT",0)
SAd=lookup("45shipdepotlookup","depot_code",«Depot»,"MAd","ADDRESS",0)
Cit=lookup("45shipdepotlookup","depot_code",«Depot»,"City","CITY",0)
Sta=lookup("45shipdepotlookup","depot_code",«Depot»,"St","ST",0)
Telephone=lookup("45shipdepotlookup","depot_code",«Depot»,"phone","0",0)
Z=lookup("45shipdepotlookup","depot_code",«Depot»,"Zip","",0)
ShipmentDescription="Farm Supplies"
GoForm "truck label"
PrintOneRecord Dialog
case «TruckingCo» contains "land"
GoForm "landairBOL"
PrintOneRecord Dialog
case «TruckingCo» contains "ross"
GoForm "rossbol"
PrintOneRecord Dialog
defaultcase
GoForm "genericBOL"
PrintOneRecord Dialog
endcase
gosheet
GetScrap "Enter PRO Number" 
«UPSTracking#»=clipboard()
call .updateshipping
window "45shipping"
Contact=title
window "45ogstally"
removeallsummaries
goform "ogspagecheck"
selectall
___ ENDPROCEDURE depot _________________________________________________________

___ PROCEDURE Groups&Depots ____________________________________________________
local vwhat
global vcode
waswindow=info("windowname")
window "Groups&Depots"
if error
openfile "Groups&Depots"
endif
deleteall
window waswindow
neworder=""
order=""
yesno "Do You Already Have a Selection Made?"
case clipboard() contains "y"
goto skip
defaultcase
getscrap "Which Depot or Group?"
if striptoalpha(clipboard())≠""
vwhat="depot"
vcode=clipboard()
select ShipCode="J" and Depot contains clipboard() and Status≠"Com"
if info("selected")=info("records")
message "This depot is done"
stop
endif
else
select int(OrderNo)=int(val(clipboard()))
endif
selectwithin OrderNo<600000
endcase
skip:
firstrecord
loop
    order=Order
    ArrayFilter order,neworder,¶,str(OrderNo)+¬+str(Sequence)+¬+import()
    window "Groups&Depots"
    openfile "+@neworder"
    window waswindow
    downrecord
    order=""
    neworder=""
until info("stopped")
window "Groups&Depots"
field Item
sortup
select Description=""
removeselected
select sz≥30
goform "To Pull"
print dialog
selectreverse
print dialog
gosheet
window waswindow
if vwhat="depot"
yesno "Do potatoes?"
if clipboard() contains "y"
call "Groups&DepotsMoose"
endif
endif

___ ENDPROCEDURE Groups&Depots _________________________________________________

___ PROCEDURE depotcheck _______________________________________________________
local wait, number, veepot
gettext "Which Depot?", veepot
yesno "Include Spuds?"
if clipboard() contains "y"
select Status≠"Com" and OrderNo=int(OrderNo) and Depot=veepot
else
select Status≠"Com" and OrderNo=int(OrderNo) and OrderNo<600000 and Depot=veepot
endif
field ShippingWt
Total
field OrderNo
Count
lastrecord
message str(ShippingWt)+" lbs"+¶+str(OrderNo)+" orders"
removeallsummaries

___ ENDPROCEDURE depotcheck ____________________________________________________

___ PROCEDURE dropship _________________________________________________________
local twonotes
global vvoice
vvoice=OrderNo
Field «Notes2»
twonotes=«Notes2»
«Notes2»=twonotes+" dropship order placed "+datepattern(today(),"mm-dd-yy")
openfile "45ogsdropship"
openfile "NewSupplier"
window "45ogscomments.linked"
if error
window "45ogscomments.linked:real for ordering"
else
goform "real for ordering"
endif
___ ENDPROCEDURE dropship ______________________________________________________

___ PROCEDURE matchmaking ______________________________________________________
local supplyorder, spudorder, person, notice
select OrderNo<400000 and OrderNo > 300000 and ((ShipCode="P" and Status≠"Com") or (ShipCode="P" and PickedUp <> "Yes") or ShipCode="T")
selectadditional OrderNo≥600000 and (ShipCode="P" or ShipCode="T")
selectwithin OrderNo=int(OrderNo)
field OrderNo
sortup
firstrecord
loop
person=«C#»
supplyorder=str(OrderNo)
find OrderNo>600000 and «C#»=person
if info("found")
    spudorder=str(OrderNo)
    notice=OrderComments
    OrderComments=notice+" also has "+supplyorder
    find OrderNo=val(supplyorder)
    notice=OrderComments
    OrderComments=notice+" also has "+spudorder
    downrecord
else
    downrecord
endif
until OrderNo≥600000
___ ENDPROCEDURE matchmaking ___________________________________________________

___ PROCEDURE check continuity _________________________________________________
field OrderNo
sortup
;firstrecord
local numero
loop
numero=OrderNo
downrecord
case OrderNo≠int(OrderNo)
    loop
    downrecord
    until OrderNo=int(OrderNo)
    if OrderNo≠numero+1
    message "Missing an order"
    stop
    endif
case OrderNo≠numero+1
message "Missing an order"
stop
endcase
stoploopif OrderNo≥310000
while forever
___ ENDPROCEDURE check continuity ______________________________________________

___ PROCEDURE Tree Sale ________________________________________________________
waswindow=info("windowname")
window "Groups&Depots"
if error
openfile "Groups&Depots"
endif
deleteall
window waswindow
neworder=""
order=""
select ShipCode="T" and OrderNo<600000
firstrecord
loop
    order=Order
    ArrayFilter order,neworder,¶,¬+str(OrderNo)+¬+str(Sequence)+¬+import()
    window "Groups&Depots"
    openfile "+@neworder"
    window waswindow
    downrecord
    order=""
    neworder=""
stoploopif OrderNo=600001
until info("stopped")
window "Groups&Depots"
field Item
sortup
group


___ ENDPROCEDURE Tree Sale _____________________________________________________

___ PROCEDURE textport/† _______________________________________________________
waswindow=info("windowname")
case info("formname")="ogspagecheck"
    getscrap "last OGS order"
    Numb=val(clipboard())
    select OrderNo > 300000 and OrderNo≤Numb
case info("formname")="mtpagecheck"
    getscrap "last MT order"
    Numb=val(clipboard())
    select OrderNo≤Numb and OrderNo≥600000
case info("formname")="bulbspagecheck"
    getscrap "last Bulbs order"
    Numb=val(clipboard())
    select OrderNo≤Numb and OrderNo≥500000
endcase

selectwithin OrderNo=int(OrderNo)
openfile "Textport"
openfile "&&" + "45ogstally"
call "textport"
selectall
find OrderNo=Numb
save
___ ENDPROCEDURE textport/† ____________________________________________________

___ PROCEDURE soiltest2 ________________________________________________________
waswindow=info("windowname")
local order,neworder,vnum
order=""
neworder=""
vnum=0
openfile "soiltests"
gosheet
deleteall
window waswindow
selectall
select
Order contains "8194" and «Notes1» notcontains "soil test" and Status="" 
if info("selected")=info("records")
message "All taken care of"
stop
else
field «Notes1»
formulafill
«Notes1»+¶+"soil test printed"
endif

loop
order=Order
arrayfilter order,neworder,¶,str(OrderNo)+¬+Con+¬+MAd+¬+City+¬+St+¬+str(Zip)+¬+import()
window "soiltests"
openfile "+@neworder"
window waswindow
downrecord
order=""
neworder=""
until info("stopped")

window "soiltests"
select Item contains "8194"
removeunselected
firstrecord



window "soiltests"
local vnum,enn
vnum=0
loop
vnum=val(qty)
enn=1
    loop
    copyrecord
    pasterecord
    until (vnum*2)-1
    loop
    OrderNo=OrderNo+"-"+str(enn)
    downrecord
    OrderNo=OrderNo+"-"+str(enn)
    MAd="PO Box 520"
    City="Clinton"
    St="ME"
    Zip=04927
    enn=enn+1
    downrecord
    until vnum
until info("stopped")
firstrecord 
call "add blanks"
    
goform "label"
print dialog
___ ENDPROCEDURE soiltest2 _____________________________________________________

___ PROCEDURE fillbackorder ____________________________________________________
waswindow=info("windowname")
global vedad,vrate
vedad=TaxState
vrate=TaxRate
fileglobal linenu, newcomment,sel, vord, sub
case info("formname")="ogspagecheck"
    loop
        GetScrap "Search for what item?"
        Select BackOrder Contains str(clipboard())
        stoploopif info("empty")

        repeat:

        firstrecord
        loop
            linenu=0
            linenu=arraysearch(BackOrder, "*"+str(clipboard())+"*",1,¶)
                if linenu=0
                    goto nextrecord
                endif
                ;message extract(BackOrder, ¶, linenu)
                If extract(BackOrder, ¶, linenu) contains "backorder"
                    BackOrder=
                    arraychange(BackOrder,replace(extract(BackOrder,¶,linenu),"backorder",rep(chr(32),9)),linenu,¶)
                    BackOrder=
                    arraychange(BackOrder,arraychange(array(BackOrder,linenu,¶), rep(chr(95),4), 5,¬),linenu,¶)
                    else
                    BackOrder=BackOrder
                    endif
            linenu=arraysearch(Order, "*"+str(clipboard())+"*",1,¶)
            Order=
            arraychange(Order,replace(extract(Order,¶,linenu),"backorder",rep(chr(32),9)),linenu,¶)
            Order=
            arraychange(Order,arraychange(array(Order,linenu,¶), extract(array(Order,linenu,¶),¬,4), 5,¬),linenu,¶)
            nextrecord:
            downrecord
        until info("stopped")

        YesNo "Do you want to fill another backorder?"
            if clipboard()="Yes"
                sel=info("selected")
                GetScrap "Search for what item?"
                SelectAdditional Backorder contains str(clipboard())
                stoploopif info("selected")=sel
                GoTo repeat
            endif
        stoploopif clipboard()="No"
    until forever
                                                          
    selectwithin BackOrder notcontains "backorder"
    if info("empty")
        message "there are no orders ready to fill"
        stop
    endif
    
    stop
    openfile "NewOGSTotaller"
    window "45ogscomments.linked:secret"
    ReSynchronize
    window "45ogscomments:secret"
    openfile "&&" + "45ogscomments.linked"
    window waswindow

    firstrecord
    loop
        order=BackOrder
        window "NewOGSTotaller:secret"
        openfile "&@order"
        call ".backorders"
        downrecord
    until info("stopped")

    openform "backorder"
    print dialog
    closewindow
    selectall
 case info("formname") ="mtpagecheck"
            openfile "45mt prices"
            window waswindow
            select Backorder≠"" and OrderNo>600000 
            YesNo "select a range?"
                If clipboard()="Yes"
                    GetScrap "Last order to print"
                    Selectwithin OrderNo≤val(clipboard())
                endif
            Selectwithin Order contains "ship"
            field Sequence
            sortup
            firstrecord
            loop
                vord=OrderNo
                sub=Subs
                order=Order
                openfile "MT Totaller"
                openfile "&@order"
                call ".backorders"
                downrecord
            until info("stopped")
        openform "backorder"
        print ""
        closewindow
        selectall
endcase
___ ENDPROCEDURE fillbackorder _________________________________________________

___ PROCEDURE .giftinvoice _____________________________________________________
global order, waswindow
waswindow=info("windowname")
order=Order
window "NewOGSTotaller"
openfile "&@order"
call .giftinvoice

___ ENDPROCEDURE .giftinvoice __________________________________________________

___ PROCEDURE Groups&DepotsMoose _______________________________________________
local nopull
nopull="7490,7500,7517,7519,7520,7545,7550,7990,7995,7997,7998,7999"
waswindow=info("windowname")
window "Groups&DepotsMoose"
if error
openfile "Groups&DepotsMoose"
endif
deleteall
window waswindow
selectall
neworder=""
order=""
yesno "Do You Already Have a Selection Made?"
case clipboard() contains "y"
goto skip
defaultcase
getscrap "Which Depot or Group?"
if striptoalpha(clipboard())≠""
vwhat="depot"
vcode=clipboard()
select ShipCode="J" and Depot contains clipboard() and Status≠"Com"
if info("selected")=info("records")
message "This depot is done"
stop
endif
else
select int(OrderNo)=int(val(clipboard()))
endif
selectwithin OrderNo≥600000
endcase
skip:
loop
    order=Order
    ArrayFilter order,neworder,¶,str(OrderNo)+¬+str(Sequence)+¬+import()
    window "Groups&DepotsMoose"
    openfile "+@neworder"
    window waswindow
    downrecord
    order=""
    neworder=""
until info("stopped")
window "Groups&DepotsMoose"
select Description=""
if info("selected")<info("records")
removeselected
endif
select comment contains "out"
if info("selected")<info("records")
removeselected
endif
select arraycontains(nopull,Item,",")=0
field SzCode
sortup
field Item
sortupwithin
goform "To Pull"
print dialog
gosheet
___ ENDPROCEDURE Groups&DepotsMoose ____________________________________________

___ PROCEDURE .schedulepickup __________________________________________________
local vpickup
gettext "When are they picking up?", vpickup
«PickUpDate»=date(vpickup)
ShipCode="P"
___ ENDPROCEDURE .schedulepickup _______________________________________________

___ PROCEDURE fedexport ________________________________________________________
local importarray, madimportarray
synchronize
selectall

;; select all OGS orders that have a non-blank status and were completed or backordered within the past 30 days
select OrderNo > 300000 and OrderNo < 400000 and ((Status <> "" and FillDate > today()-90) or OrderPlaced > today()-90)

;; some POE orders need shipping quotes, which may be generated before POE shipping begins, so these orders will always
;; be included in the export even if POE shipping hasn't started yet.
selectadditional OrderNo > 600000 and OrderNo < 700000 and Status notcontains "c" and ShippingWt >= 200 and ShippingWt <= 500

;; during the main potato season, we'll have to select all POE orders, because many of them were placed months ago
if today() >= date("February 15") and today() <= date("May 20")
    selectadditional OrderNo > 600000 and OrderNo < 700000 and (Order contains '7990' or Order contains '7995')
    if today() >= date("March 25")
        selectadditional OrderNo > 600000 and OrderNo < 700000
    endif
endif

;; ALSO select all Bulbs orders
if today() >= date("August 15") and today() <= date("November 20")
    selectadditional OrderNo > 500000 and OrderNo < 600000
endif

ono=OrderNo
importarray=""

arrayselectedbuild importarray,chr(10),"",?(Group≠"",Group+","+Con," "+","+Con)+","+SAd+","+Cit+","+Sta+","+pattern(Z,"#####")+","+Email+","+str(OrderNo)+Con[1,1]+Cit[1,1]+chr(13)
selectwithin SAd <> MAd
arrayselectedbuild madimportarray,chr(10),"",?(Group≠"",Group+",MailTo: "+Con," "+",MailTo: "+Con)+","+MAd+","+City+","+St+","+pattern(Zip,"#####")+","+Email+","+str(OrderNo)+Con[1,1]+Cit[1,1]+chr(13)

importarray="Group"+","+"Con"+","+"SAd"+","+"Cit"+","+"Sta"+","+"Z"+","+"Email"+","+"OrderNo"+Con[1,1]+Cit[1,1]+chr(13)+chr(10)+importarray+chr(10)+madimportarray

select OrderNo=ono
export "fedexport.csv", importarray
selectall

___ ENDPROCEDURE fedexport _____________________________________________________

___ PROCEDURE forcesynchronize _________________________________________________
forcesynchronize
___ ENDPROCEDURE forcesynchronize ______________________________________________

___ PROCEDURE (Bulbs) __________________________________________________________

___ ENDPROCEDURE (Bulbs) _______________________________________________________

___ PROCEDURE justbulbs/B ______________________________________________________
select OrderNo >= 500000 and OrderNo < 600000
___ ENDPROCEDURE justbulbs/B ___________________________________________________

___ PROCEDURE bulbs exporter/2 _________________________________________________
global order, wTally, vord1, vord2, foldername, parentfoldername

foldername=""
wTally = info("WindowName")
openfile "45bulbs lookup"
openfile "bulbextractor"
deleteall
window wTally
field OrderNo
sortup

// selecting bulbs order numbers only //

selectwithin OrderNo > 500000 AND OrderNo < 600000
selectwithin Order notcontains "0000" and Order ≠ ""
firstrecord

noshow

     Loop
         ;;ArrayFilter Order,order,¶,str(OrderNo)+¬+extract(Order,¶,seq())+¬+str(Sequence)
         if PickSheet = ""
            ArrayFilter Order,order,¶,str(OrderNo)+¬+arrayrange(extract(Order,¶,seq()),1,8,¬)+¬+str(Sequence)
         else
         ;; this is grabbing the "fill" quantity from the order field, not the "qty." 
         ;; If you want qty instead, change the 4 to a 3
            ArrayFilter Order,order,¶,str(OrderNo)+¬+str(val(array(extract(Order,¶,seq()),1,¬)))+¬+array(extract(Order,¶,seq()),2,¬)+¬+array(extract(Order,¶,seq()),3,¬)
            +¬+array(extract(Order,¶,seq()),5,¬)+¬+array(extract(Order,¶,seq()),6,¬)+¬+array(extract(Order,¶,seq()),7,¬)+¬+array(extract(Order,¶,seq()),8,¬)+¬+¬+str(Sequence)
         endif
         
         ;;want to see what this array does? unmute the line below
         displaydata order
         window "bulbextractor:secret"
         openfile "+@order"
             ;;this is appending the data from the "order" variable/array into the bulbextracor file
         window wTally
                 downrecord
     until info("Stopped")

endnoshow

openfile "bulbextractor"


firstrecord
if OrderNo=0
     deleterecord
endif
vord1=int(OrderNo)-500000
LastRecord
vord2=int(OrderNo)-500000

field Count
fill zeroblank(0)

save

parentfoldername=folderpath(dbinfo("folder",""))
foldername=folderpath(dbinfo("folder","")) +str(vord1)+"-"+str(vord2)

makefolder foldername


;;;;;;;;;;;;;;;;;;

Field Count
select Size = "A"
formulafill val(lookup("45bulbs lookup","number",Item,"lookup size A","",0))
select Size = "B"
formulafill val(lookup("45bulbs lookup","number",Item,"size B","",0))
select Size = "C"
formulafill val(lookup("45bulbs lookup","number",Item,"size C","",0))
select Size = "D"
formulafill val(lookup("45bulbs lookup","number",Item,"size D","",0))
selectall


Field TotalCount
formulafill Count*Qty
Field RunningTotal
formulafill TotalCount

Field "Item"
SortUp
GroupUp
Field "Sequence"
SortUp

Field RunningTotal
runningtotal
removeallsummaries


Field "CountOrderedFromSuppliers"

FormulaFill lookup("45BulbsInventory","Item",«Item»,"inventorybulk",0,0)
;;old lookup
;;lookup("bulbs purchased","Item",Item,"Qty",0,0)
Select RunningTotal <= CountOrderedFromSuppliers

Field "Item"
GroupUp

Field "Sequence"
Maximum

Field LastSeqToFill
formulafill lookup(info("databasename"),"Item",Item,"Sequence",0,1)

removeallsummaries
Selectall
Field "Item"
GroupUp

Field LastSeqToFill
maximum

Field CountOrderedFromSuppliers
maximum

Field RunningTotal
propagate

CollapseToLevel "1"
SelectWithin CountOrderedFromSuppliers > RunningTotal

Field "LastSeqToFill"
Fill 9999

SelectAll

SaveACopyAs foldername + ":45bulb running totals "+str(vord1)+"-"+str(vord2)

;;debug

Revert
removeallsummaries

;;;;;;;;;;;;;;;;;;;;;;;;
     SaveAs foldername + ":45bulb collation " +str(vord1)+"-"+str(vord2)

         field OrderNo
         GroupUp
         Field Total
         Total
         OutlineLevel "1"



         SaveACopyAs foldername + ":45bulb subtotals "+str(vord1)+"-"+str(vord2)
         Revert
         Field Size
         GroupUp
         Field Item
         GroupUp
         Field Qty
         Total
         Field Size
         Propagate
         RemoveDetail "Data"
         RemoveSummaries 1
         RemoveSummaries 2
         Field Item
         Sortup


         SaveAs foldername + ":45bulb packets "+str(vord1)+"-"+str(vord2)
             ;; look for itemtotals in the parent folder

             ;;debug

             OpenFile parentfoldername + "45bulbs itemtotals"
             OpenFile "&" + foldername + ":45bulb packets "+str(vord1)+"-"+str(vord2)

             ;; make itemtotals save in the new folder
             Call itemtotals

___ ENDPROCEDURE bulbs exporter/2 ______________________________________________

___ PROCEDURE SelectEarly-unprinted ____________________________________________
synchronize
selectall
field OrderNo
sortup

;; selects Bulbs orders with ANY early shipping items, where the orders have NOT been printed

arrayfilter earlyItems,filteredEarly,",","Order contains "+quoted(import())
filteredEarly=replace(filteredEarly,","," or ")

selectwithin PickSheet="" and OrderNo >= 500000 and OrderNo < 600000
execute " selectwithin "+filteredEarly

if info("empty")
    message "nothing selected"
endif
___ ENDPROCEDURE SelectEarly-unprinted _________________________________________

___ PROCEDURE SelectEarly-printed ______________________________________________
synchronize
selectall
field OrderNo
sortup

;; selects Bulbs orders with ANY early shipping items, where the orders have already been printed

arrayfilter earlyItems,filteredEarly,",","Order contains "+quoted(import())
filteredEarly=replace(filteredEarly,","," or ")

selectwithin PickSheet≠"" and OrderNo >= 500000 and OrderNo < 600000
execute " selectwithin "+filteredEarly

if info("empty")
    message "nothing selected"
endif
___ ENDPROCEDURE SelectEarly-printed ___________________________________________

___ PROCEDURE ----- ____________________________________________________________

___ ENDPROCEDURE ----- _________________________________________________________

___ PROCEDURE SelectMain-unprinted _____________________________________________
; This query selects all bulbs orders containing ONLY early or ONLY late items, then reverses the selection to get orders with main shipment items.

arrayfilter lateItems,filteredLate,",","import() contains "+quoted(import())
filteredLate=replace(filteredLate,","," or ")
arrayfilter earlyItems,filteredEarly,",","import() contains "+quoted(import())
filteredEarly=replace(filteredEarly,","," or ")

selectall
select OrderNo >= 500000 and OrderNo < 600000 and ShipCode≠"D" and Order <>"" and Status notcontains "c" and Order NOTCONTAINS "1)"+¬+"0"+¬+¬+"0"
selectwithin arraystrip(arrayfilter(replace(Order,lf(),cr()),cr(),{?(}+filteredLate+{,"",import())}),cr())="" or arraystrip(arrayfilter(replace(Order,lf(),cr()),cr(),{?(}+filteredEarly+{,"",import())}),cr())=""
selectreverse
selectwithin OrderNo >= 500000 and OrderNo < 600000 and ShipCode≠"D" and Order <>"" and Status notcontains "c" and Order NOTCONTAINS "1)"+¬+"0"+¬+¬+"0"

___ ENDPROCEDURE SelectMain-unprinted __________________________________________

___ PROCEDURE - ________________________________________________________________

___ ENDPROCEDURE - _____________________________________________________________

___ PROCEDURE SelectLate-unprinted _____________________________________________
synchronize
selectall
field OrderNo
sortup
;; selects Bulbs orders with ANY late shipping items, where the orders have NOT been printed

arrayfilter lateItems,filteredLate,",","Order contains "+quoted(import())
filteredLate=replace(filteredLate,","," or ")

selectwithin PickSheet="" and OrderNo >= 500000 and OrderNo < 600000
execute " selectwithin "+filteredLate

if info("empty")
    message "nothing selected"
endif
___ ENDPROCEDURE SelectLate-unprinted __________________________________________

___ PROCEDURE SelectLate-printed _______________________________________________
synchronize
selectall
field OrderNo
sortup

;; selects Bulbs orders with ANY late shipping items, where the orders have already been printed

arrayfilter lateItems,filteredLate,",","Order contains "+quoted(import())
filteredLate=replace(filteredLate,","," or ")

selectall
selectwithin PickSheet≠"" and OrderNo >= 500000 and OrderNo < 600000
execute " selectwithin "+filteredLate

if info("empty")
    message "nothing selected"
endif

___ ENDPROCEDURE SelectLate-printed ____________________________________________

___ PROCEDURE SelectLate-only_orders ___________________________________________
message "this makes a selection of late shipment only orders to be flagged by changing the ship code to L" 

arrayfilter lateItems,filteredLate,",","import() contains "+quoted(import())
filteredLate=replace(filteredLate,","," or ")

selectall
selectwithin ShipCode≠"D" and Order <>"" and Status notcontains "c" and Order NOTCONTAINS "1)"+¬+"0"+¬+¬+"0" and OrderNo >= 500000 and OrderNo < 600000
selectwithin arraystrip(arrayfilter(replace(Order,lf(),cr()),cr(),{?(}+filteredLate+{,"",import())}),cr())=""

if info("empty")
    message "there are no orders with JUST late-shipping items"
endif
___ ENDPROCEDURE SelectLate-only_orders ________________________________________

___ PROCEDURE -- _______________________________________________________________

___ ENDPROCEDURE -- ____________________________________________________________

___ PROCEDURE SelEarlyOrdersToPrint ____________________________________________
local selectitems, scratchearlyitems, query, waswindow, temp_selected_items

waswindow = info("windowname")
selectitems=""

window "45BulbsComments linked"

select section contains "early ship" and «ok to ship» contains "A"
if info("empty")
else
    arrayselectedbuild temp_selected_items,",","45BulbsComments linked",str(number)+¬+"A"
    selectitems = temp_selected_items
endif

select section contains "early ship" and «ok to ship» contains "B"
if info("empty")
else
    arrayselectedbuild temp_selected_items,",","45BulbsComments linked",str(number)+¬+"B"
    selectitems = selectitems + "," + temp_selected_items
endif

select section contains "early ship" and «ok to ship» contains "C"
if info("empty")
else
    arrayselectedbuild temp_selected_items,",","45BulbsComments linked",str(number)+¬+"C"
    selectitems = selectitems + "," + temp_selected_items
endif

selectall

window waswindow

synchronize
selectall
field OrderNo
sortup

;; this macro is intended for early in the season when not all items have arrived
;; or been repacked yet, but we want to start filling whatever orders we can.

;; Items that are in house and ready to be pulled for orders should be marked as "ABC" 
;; (whichever size codes are ready) in the "ok to ship" field 45BulbsComments linked
;; Any items where "ok to ship" is blank will be treated as if they are NOT available to ship yet.

;; it selects Bulbs orders with JUST the entered early items (though
;; non-early items are allowed, because we won't ship those now anyway)

;; selectItems is a list of the early items that are available to ship now
;; earlyItems is a list of ALL early items in the catalog
;; scratchearlyitems is the difference between earlyItems - selectItems
;; so it's all of the early items that we DON'T have available yet.

arraydifference earlyItems, selectitems, ",", scratchearlyitems 

;; Order notcontains any of the unavailable early items
arrayfilter scratchearlyitems,scratchearlyitems,",","Order notcontains "+quoted(import())
scratchearlyitems=replace(scratchearlyitems,","," and ")

;; and Order DOES contains at least one of the entered early items
arrayfilter selectitems,selectitems,",","Order contains "+quoted(import())
selectitems=replace(selectitems,","," or ")

query=scratchearlyitems + " and (" + selectitems + ")"

select OrderNo > 500000 and OrderNo < 600000

execute "selectwithin "+query

;; only search for orders that haven't been printed yet
selectwithin Status notcontains "c" and PickSheet=""
selectwithin «in_process» notcontains "p"
___ ENDPROCEDURE SelEarlyOrdersToPrint _________________________________________

___ PROCEDURE SelRegOrdersToPrint ______________________________________________
local selectitems, scratchregularitems, query, allitems, temp_selected_items

synchronize

selectall
field OrderNo
sortup

waswindow = info("windowname")
selectitems=""

;; this macro is intended for early in the season when not all items have arrived
;; or been repacked yet, but we want to start filling whatever orders we can.

;; Items that are in house and ready to be pulled for orders should be marked as "ABC" 
;; (whichever size codes are ready) in the "ok to ship" field 45BulbsComments linked
;; Any items where "ok to ship" is blank will be treated as if they are NOT available to ship yet.

;; it selects Bulbs orders with JUST the entered REGULAR items (though
;; early and late items are allowed, because those would ship separately anyway)

window "45BulbsComments linked"

select section notcontains "early ship" and section notcontains "late" and «ok to ship» contains "A"
if info("empty")
else
    arrayselectedbuild temp_selected_items,",","45BulbsComments linked",str(number)+¬+"A"
    selectitems = temp_selected_items
endif

select section notcontains "early ship" and section notcontains "late" and «ok to ship» contains "B"
if info("empty")
else
    arrayselectedbuild temp_selected_items,",","45BulbsComments linked",str(number)+¬+"B"
    selectitems = selectitems + "," + temp_selected_items
endif

select section notcontains "early ship" and section notcontains "late" and «ok to ship» contains "C"
if info("empty")
else
    arrayselectedbuild temp_selected_items,",","45BulbsComments linked",str(number)+¬+"C"
    selectitems = selectitems + "," + temp_selected_items
endif

selectall

;; selectwithin Order contains x or Order contains y or Order contains z...

;; allitems = import from comments.linked
openfile "45BulbsComments linked"
select number > 10
arrayselectedbuild allitems, ",", "45BulbsComments linked", str(number)+¬+"A,"+str(number)+¬+"B,"+str(number)+¬+"C"

;; scratchregularitems = allitems - earlyItems - lateItems - selectitems
arraydifference allitems, earlyItems, ",", scratchregularitems
arraydifference scratchregularitems, lateItems, ",", scratchregularitems
arraydifference scratchregularitems, selectitems, ",", scratchregularitems

window waswindow

;; Order notcontains any of the unavailable regular items
arrayfilter scratchregularitems,scratchregularitems,",","Order notcontains "+quoted(import())
scratchregularitems=replace(scratchregularitems,","," and ")


;; and Order DOES contains at least one of the entered regular items
arrayfilter selectitems,selectitems,",","Order contains "+quoted(import())
selectitems=replace(selectitems,","," or ")


query=scratchregularitems + " and (" + selectitems + ")"

select OrderNo > 500000 and OrderNo < 600000

execute "selectwithin "+query
selectwithin Status notcontains "c" and Order NOTCONTAINS "1)"+¬+"0"+¬+¬+"0"
selectwithin «in_process» notcontains "p"
___ ENDPROCEDURE SelRegOrdersToPrint ___________________________________________

___ PROCEDURE SelOrdersToPrint _________________________________________________
local selectitems, scratchregularitems, query, allitems, temp_selected_items

synchronize

selectall
field OrderNo
sortup

waswindow = info("windowname")
selectitems=""

;; this macro is intended to search for orders that can be filled *regardless*
;; of whether the items are early/regular/late season shippers

;; Items that are in house and ready to be pulled for orders should be marked as "ABC" 
;; (whichever size codes are ready) in the "ok to ship" field 45BulbsComments linked
;; Any items where "ok to ship" is blank will be treated as if they are NOT available to ship yet.

;; it selects Bulbs orders with JUST the entered REGULAR items (though
;; early and late items are allowed, because those would ship separately anyway)

window "45BulbsComments linked"

select «ok to ship» contains "A"
if info("empty")
else
    arrayselectedbuild temp_selected_items,",","45BulbsComments linked",str(number)+¬+"A"
    selectitems = temp_selected_items
endif

select «ok to ship» contains "B"
if info("empty")
else
    arrayselectedbuild temp_selected_items,",","45BulbsComments linked",str(number)+¬+"B"
    selectitems = selectitems + "," + temp_selected_items
endif

select «ok to ship» contains "C"
if info("empty")
else
    arrayselectedbuild temp_selected_items,",","45BulbsComments linked",str(number)+¬+"C"
    selectitems = selectitems + "," + temp_selected_items
endif

selectall



;; allitems = import from comments.linked
openfile "45BulbsComments linked"
select number > 10
arrayselectedbuild allitems, ",", "45BulbsComments linked", str(number)+¬+"A,"+str(number)+¬+"B,"+str(number)+¬+"C"

;; scratchregularitems = allitems - selectitems
arraydifference allitems, selectitems, ",", scratchregularitems

window waswindow

;; Order notcontains any of the unavailable regular items
arrayfilter scratchregularitems,scratchregularitems,",","Order notcontains "+quoted(import())
scratchregularitems=replace(scratchregularitems,","," and ")


query=scratchregularitems

select OrderNo > 500000 and OrderNo < 600000

execute "selectwithin "+query
selectwithin Status notcontains "c" and Order NOTCONTAINS "1)"+¬+"0"+¬+¬+"0"
selectwithin OrderNo = int(OrderNo)
selectwithin GroupMarker = ''
selectwithin «in_process» notcontains "p"
___ ENDPROCEDURE SelOrdersToPrint ______________________________________________

___ PROCEDURE Select_unrun_paperwork ___________________________________________
synchronize
selectall
field OrderNo
sortup

select «in_process»≠"" and OrderNo >= 500000 and OrderNo < 600000

if info("selected")=info("records")
message "nothing selected"
endif
___ ENDPROCEDURE Select_unrun_paperwork ________________________________________

___ PROCEDURE .complete ________________________________________________________
if info("files") notcontains "45ogstaxable"
message "tax file not open"
endif
___ ENDPROCEDURE .complete _____________________________________________________

___ PROCEDURE .NOFA_settle_up __________________________________________________
local nofa_original_paid, nofa_total_freight, nofa_difference
global whodis

select OrderNo > 600000 and Depot contains "NOFA"
adjusting_order = 1

field AddPay2
Total
if AddPay2 <> 0
    message "The AddPay2 field is already used. Fix it before running this macro."
    stop
endif
RemoveAllSummaries

gettext "Total freight charge to ship all POE NOFA?", nofa_total_freight

field GrTotal
Total
nofa_original_paid = GrTotal
RemoveAllSummaries

Field Notes2
formulafill Notes2 + " customer paid: " + pattern(GrTotal,"$#,.##")

Field «$Shipping»
formulafill val(nofa_total_freight)/info("selected")

Field ShipCode
fill "C"

;; loop through orders, mark all as bulk and adjust them
Field Bulk
fill "bulk"

openform "mtpagecheck"

debug

firstrecord
loop
    whodis = "NOFA_settle_up"
    call "change/h"
    ;; need to bypass the "next"
    downrecord
until info("eof")

;;;;;;;;;;;;;

field «BalDue/Refund»
Total
nofa_difference = «BalDue/Refund»
RemoveAllSummaries

field AddPay2
formulafill -«BalDue/Refund»

field DatePay2
formulafill today()

field MethodPay2
fill "Check"

field «BalDue/Refund»
fill 0

field «1SpareText» 
formulafill datepattern(today(),"mm/dd/yy") + " " + timepattern(now(),"HH:MM AM/PM")

message "NOFA customers originally paid " + pattern(nofa_original_paid,"$#,.##") + ". Send NOFA a check for " + pattern(nofa_difference,"$#,.##")
___ ENDPROCEDURE .NOFA_settle_up _______________________________________________

___ PROCEDURE .temp_export_to_landscape ________________________________________
local vresponse, vtoken, vfolderpath, vfilename, vfieldnames, varray, mycurl

varray=""
vfilename="orderrecords.csv"
;; setservervariable "", landscapeauthtoken, "1|Cn9qoS8srEsHVte8js5Mg9odE4EW6VdQat9m5GAM" ;; can be commented out after the real token has been set
vtoken = servervariable("","landscapeauthtoken")
vfolderpath = "/" + arraydelete(replace(dbpath(),":","/"),1,1,"/")

vfieldnames = replace(dbinfo("fields",""),¶,",")
arrayfilter vfieldnames, vfieldnames, ",",'"'+import()+'"'
    
hide
;; C# is the first field in the database
field «C#»
varray = str(«»)
right
loop
    if info("datatype") < 5
        varray = varray + ',"' + str(«») + '"'
    else
        varray = varray + ',' + str(«»)
    endif
    right
until info("stopped")
show

field OrderNo

filesave "", vfilename, "", vfieldnames+¶
fileappend "", vfilename,  "", replace(varray,cr(),lf())

mycurl = '-F computername="' + info("computername") + '" -F csv=@"' + vfolderpath + vfilename + '" -H "Authorization: Bearer ' + vtoken + '" http://localhost/api/v1/panorama-csv'

;;This emulates a filled-in form in which a user has pressed the submit button. 
;; When prefixing the filename with @, the file is attached to the post as a file upload. 
;; The other value submitted (computername) shows up as a multipart value.

curl vresponse, mycurl

;; this line is just here for debugging--can be removed when this code is in use
displaydata vresponse




___ ENDPROCEDURE .temp_export_to_landscape _____________________________________

___ PROCEDURE .test ____________________________________________________________
messsage arraysize(extract(OGSOrder,¶,1),¬)
___ ENDPROCEDURE .test _________________________________________________________

___ PROCEDURE (CommonFunctions) ________________________________________________

___ ENDPROCEDURE (CommonFunctions) _____________________________________________

___ PROCEDURE ExportMacros _____________________________________________________
local Dictionary1, ProcedureList
//this saves your procedures into a variable
exportallprocedures "", Dictionary1
clipboard()=Dictionary1

message "Macros are saved to your clipboard!"
___ ENDPROCEDURE ExportMacros __________________________________________________

___ PROCEDURE ImportMacros _____________________________________________________
local Dictionary1,Dictionary2, ProcedureList
Dictionary1=""
Dictionary1=clipboard()
yesno "Press yes to import all macros from clipboard"
if clipboard()="No"
stop
endif
//step one
importdictprocedures Dictionary1, Dictionary2
//changes the easy to read macros into a panorama readable file

 
//step 2
//this lets you load your changes back in from an editor and put them in
//copy your changed full procedure list back to your clipboard
//now comment out from step one to step 2
//run the procedure one step at a time to load the new list on your clipboard back in
//Dictionary2=clipboard()
loadallprocedures Dictionary2,ProcedureList
message ProcedureList //messages which procedures got changed

___ ENDPROCEDURE ImportMacros __________________________________________________

___ PROCEDURE Symbol Reference _________________________________________________
bigmessage "Option+7= ¶  [in some functions use chr(13)
Option+= ≠ [not equal to]
Option+\= « || Option+Shift+\= » [chevron]
Option+L= ¬ [tab]
Option+Z= Ω [lineitem or Omega]
Option+V= √ [checkmark]
Option+M= µ [nano]
Option+<or>= ≤or≥ [than or equal to]"


___ ENDPROCEDURE Symbol Reference ______________________________________________

___ PROCEDURE GetDBInfo ________________________________________________________
local DBChoice, vAnswer1, vClipHold

Message "This Procedure will give you the names of Fields, procedures, etc in the Database"
//The spaces are to make it look nicer on the text box
DBChoice="fields
forms
procedures
permanent
folder
level
autosave
fileglobals
filevariables
fieldtypes
records
selected
changes"
superchoicedialog DBChoice,vAnswer1,“caption="What Info Would You Like?"
captionheight=1”


vClipHold=dbinfo(vAnswer1,"")
bigmessage "Your clipboard now has the name(s) of "+str(vAnswer1)+"(s)"+¶+
"Preview: "+¶+str(vClipHold)
Clipboard()=vClipHold

___ ENDPROCEDURE GetDBInfo _____________________________________________________

___ PROCEDURE .AutomaticFY _____________________________________________________

    
global dateHold, dateMath, intYear, 
thisFYear,lastFYear,nextFYear,intMonth,fileDate

fileDate=val(striptonum(info("databasename")))
nextFYear=""
thisFYear=""
lastFYear=""

//get the date
dateHold = datepattern(today(),"mm/yyyy")

//gets the current month and year
intMonth = val(dateHold[1,"/"][1,-2])
intYear = val(dateHold["/",-1][2,-1])

//assigns FY numbers for years

case val(intMonth)>6
    nextFYear=str(intYear-1976)
    thisFYear=str(intYear-1977)
    lastFYear=str(intYear-1978)

case val(intMonth)<7
    nextFYear=str(intYear-1977)
    thisFYear=str(intYear-1978)
    lastFYear=str(intYear-1979)

endcase

//checks if this is an older file and needs older FYs
if fileDate ≤ val(lastFYear) and fileDate > 0
    nextFYear=str(fileDate+1)
    thisFYear=str(fileDate)
    lastFYear=str(fileDate-1)
endif

//tallmessage str(nextFYear)+¬+str(thisFYear)+¬+str(lastFYear)


/*

///////~~~~~~~
Programmer Notes
~~~~~~~~~//////////
The danger of this procedure is that come July 1st of the year, it will automatically set
to open the newest files of a non-numbered Panorama file. And if those don't exist, you're 
gonna see errors. Also, a non numbered Panorama file that needs to call older files shouldn't
use this macro



To use these variables please note the following Panorama syntax rules:


filenames using variables:
    can just concatenate as a string
    
    ex:
        
    openfile str(variable)+"filename" 


field calls using variables:
    best to be only one variable and nothing else
    must be surrounded by ( )
    
    ex:
    
    field (VariableFieldName)
    
do your math and/or concatenation into the variable before calling it
    VariableFieldName=str(variable)+"fieldname"
 
field (str(variable)+"fieldname") will work but can cause errors
    
for assignments to that variable'd field 
    use «» for "current field/current cell" 
    
    ex: 
   
    «» = "10"
  
    
*/

___ ENDPROCEDURE .AutomaticFY __________________________________________________

___ PROCEDURE Folders&FilesMacros ______________________________________________

//message "This Function is meant to get you information about the folders and path your files are in for Panorama"


global commList, commWanted, clipHoldComm, buttonChoice, numChoice

commList=""
commWanted=""
clipHoldComm=""
buttonChoice=""
numChoice=0

commList=¶+
    "1 - Copy Text of folderpath"
    +¶+¬+¬+¬+¬+¬+¬+
    "1 code -- folderpath(folder(""))"
    +¶+" "+¶+
    "2 - Copy list of All Files and Folders in this folder" 
    +¶+¬+¬+¬+¬+¬+¬+
    "2 code -- listfiles(folder(""),"")"
    +¶+" "+¶+
    "3 - Copy list of All Panorama files in this folder" 
    +¶+¬+¬+¬+¬+¬+¬+
    '3 code -- listfiles(folder(""),"????KASX")'
    +¶+" "+¶+
    "4 - Copy list of All Text files in this folder" 
    +¶+¬+¬+¬+¬+¬+¬+
    '4 code -- listfiles(folder(""),"TEXT????")'

/*

//NOTE: these quotation marks “” vs "" are called smart quotes
//you get them with opt+[ and opt+shift+[
//normally for superchoicedialogs, i would use curly brackets around title or caption
//but to have this be able to be written into new files from another macro, I had
//to use smart quotes

*/
superchoicedialog commList, commWanted, 
“Title="Get File/Folder/Path"
    Caption="1 - Copy ~~~~~~ gets you the data
        1 - Code ~~~~~~ gets you the formula"
    captionheight="2"
    buttons="Ok;Cancel"
    width="800"
    height="800"”
    

        clipHoldComm=commWanted
        numChoice=striptonum(clipHoldComm)[1,3]


if commWanted[1,12] notcontains "code"

    case numChoice="1"
        tallmessage "clipboard now has: "+¶+folderpath(folder(""))
        clipboard()=folderpath(folder(""))

    case numChoice="2"
        tallmessage "clipboard now has: "+¶+listfiles(folder(""),"")
        clipboard()=listfiles(folder(""),"")
    
    case numChoice="3"
        tallmessage "clipboard now has: "+¶+listfiles(folder(""),"????KASX")
        clipboard()=listfiles(folder(""),"????KASX")

    case numChoice="4"
        tallmessage "clipboard now has: "+¶+listfiles(folder(""),"TEXT????")
        clipboard()=listfiles(folder(""),"TEXT????")

    endcase
endif

if commWanted[1,12] contains "code"
    case numChoice="1"
    clipboard()='folderpath(folder(""))'
    tallmessage "clipboard now has: "+¶+'folderpath(folder(""))'

    case numChoice="2"
    clipboard()='listfiles(folder(""),"")'
    tallmessage "clipboard now has: "+¶+'listfiles(folder(""),"")'
    
    case numChoice="3"
        tallmessage "clipboard now has: "+¶+'listfiles(folder(""),"????KASX")'
        clipboard()='listfiles(folder(""),"????KASX")'

    case numChoice="4"
        tallmessage "clipboard now has: "+¶+'listfiles(folder(""),"TEXT????")'
        clipboard()='listfiles(folder(""),"TEXT????")'

    endcase
endif
    


___ ENDPROCEDURE Folders&FilesMacros ___________________________________________

___ PROCEDURE DesignSheetExportImport __________________________________________

global vdictionary, 
name, value, ImportExportChoicelist,
fileList,choiceMade,winChoice1,winChoice2,vOptions

/*
programmer's notes

i was testing using a variable for options as stated in the reference. it seems to work with vOptions below

also tested using a call to listfiles vs putting listfiles in a variable. both seem to work

as seen in other procedures in this file instead of using curly braces we are using smartquotes because "setproceduretext" for the AddMacros fuction wont work otherwise

also, options for superchoices and other customizable dialogs are very particular in 
their syntax

caption = "dafsdf" will allow the code to run, but will not show a caption
caption="dafsdf" will actually show the caption
*/


vOptions=“caption="Choose file to export Design Sheet from"”
choiceMade=""
fileList=listfiles(folder(""),"????KASX")


superchoicedialog fileList, choiceMade, vOptions

winChoice1=choiceMade

superchoicedialog fileList, choiceMade,
“caption="Choose file to export Design Sheet to"”

winChoice2=choiceMade

window (winChoice1)
    opendesignsheet
    vdictionary=""
    firstrecord

        loop
            setdictionaryvalue vdictionary, «Field Name», «Equation»
            downrecord
        until info("stopped")

window (winChoice2)
    opendesignsheet
    firstrecord

        loop
            field «Equation»
            «» = getdictionaryvalue(vdictionary, «Field Name»)
            downrecord
        until info("stopped")


___ ENDPROCEDURE DesignSheetExportImport _______________________________________

___ PROCEDURE .FileChecker _____________________________________________________
///____________________________________________________________________________________________________________________________________
///____________________________________________________________________________________________________________________________________
///________________________________This is the .FileChecker macro in GetMacros_________________________________________________________
///____________________________________________________________________________________________________________________________________
///____________________________________________________________________________________________________________________________________


local fileNeeded,folderArray,smallFolderArray,sizeCheck,
procList,sizeCheck,procNames,procDBs,mostRecentProc

///________________________EDITME_____________
//replace this with whatever file you're error checking
//----------------------//
fileNeeded="members"    //
//----------------------//


////_____Got the file, but it's not open?_______________
case info("files") notcontains fileNeeded and listfiles(folder(""),"????KASX") contains fileNeeded
openfile fileNeeded

///________Don't got the file?__________________
case listfiles(folder(""),"????KASX") notcontains fileNeeded


    procList=arraystrip(info("procedurestack"),¶)
    sizeCheck=arraysize(procList,¶)
        if sizeCheck>1
            procList=arrayrange(procList,2,sizeCheck,¶) //this is to exclude getting recursive info about this macro, especially while testing
        else
            procList=arraystrip(info("procedurestack"),¶)
        endif

    procNames=arraycolumn(procList,1,¶,¬)
    procDBs=arraycolumn(procList,2,¶,¬)
    mostRecentProc=array(procNames,1,¶) 
    folderArray=folderpath(folder(""))
    sizeCheck=arraysize(folderArray,":")
    smallFolderArray=arrayrange(folderArray,4,sizeCheck,":")

displaydata "Error:"
+¶+
"You are missing the '"+fileNeeded+
"' Panorama file in this folder 
and can't continue the '"+mostRecentProc+"' procedure without it. 
Please move a copy of '"+fileNeeded+
"' to the appropriate folder and try the procedure again"
+¶+¶+¶+
"folder you're currently running from is: "
+¶+
smallFolderArray
+¶+¶+¶+
"current Pan files in that folder are: "
+¶+
listfiles(folder(""),"????KASX")
+¶+¶+¶+
"Pressing 'Ok' will open the Finder to your current folder"
+¶+¶+
"Press 'Stop' will stop this procedure", “title="Missing File!!!!" captionwidth=900 size=17 height=500 width=800”
//___________________________
//note, the above are "smart quotes" option+[ and option+shift+[ 
//you can also use curley braces, but another program I run will break
//if this file has thos
//___________________________

revealinfinder folder(""),""
stop

///_______File is open, but not active?______
defaultcase
window fileNeeded

endcase

/*
Example:

You are missing the 'members' Panorama file in this folder 
and can't continue this procedure without it. Please move a copy of
'members' to the appropriate folder and try the procedure again


folder you're currently running from is: 
Desktop:Panorama:FY45 Panorama Projects:GetMacros:


current Pan files in that folder are: 
GetMacros
GetMacrosDL
GetMacros44


Pressing 'Ok' will open the Finder to your current folder

Press 'Stop' will stop this procedure
*/
___ ENDPROCEDURE .FileChecker __________________________________________________

___ PROCEDURE .GetErrorLog _____________________________________________________
///____________________________________________________________________________________________________________________________________
///____________________________________________________________________________________________________________________________________
///________________________________This is the .GetErrorLog macro in GetMacros_________________________________________________________
///____________________________________________________________________________________________________________________________________
///____________________________________________________________________________________________________________________________________
/*

This can be called to with a parameter of the 
info("error") statement to display the error, give
the user the opportunity to try again or continue 
despite the error.

Either way, it makes a log of the error and what procedures,
windows, files, and variables were in use. 

-Lunar 8-22

Syntax to call:

        if error
            call .GetErrorLog,info("error")
        endif

*/
///____________________________________________________________________________________________________________________________________
///____________________________________________________________________________________________________________________________________

fileglobal fileNeeded,folderArray,smallFolderArray,sizeCheck, procList, mostRecentProc, 
panFilesList,activeFiles,allvariables,procNames,procDBs,errorList, procText, procTextArray,
lineNum, procCount, usedvariables,printVariables,strippedText,getError,errorMsg,vDb,vProc,
activeWindows,DictNameToday

//this is to keep a log of the errors
permanent errorDictionary
    errorDictionary=errorDictionary
    if error
    errorDictionary=""
    endif

errorMsg=""

getError=str(parameter(1))
    if error //if there's no parameter given, or if info("error") is blank, then say "Unknown"
    getError="Unknown"
    endif

procList=arraystrip(info("procedurestack"),¶)
    if procList="" //sometimes, there's no info in the procedure stack, and this macro shoudl stop at this point
    message "Procedure Stack is Empty -L"
    stop
    endif
sizeCheck=arraysize(procList,¶)
    if sizeCheck>1
    procList=arrayrange(procList,2,sizeCheck,¶) //this is to exclude getting recursive info about this macro, especially while testing
    else
    procList=arraystrip(info("procedurestack"),¶)
    endif

procNames=arraycolumn(procList,1,¶,¬)
procDBs=arraycolumn(procList,2,¶,¬)
mostRecentProc=array(procNames,1,¶) 
folderArray=folderpath(folder(""))

///____________more readable filepath________________
;sizeCheck=arraysize(folderArray,":")
;smallFolderArray=arrayrange(folderArray,4,sizeCheck,":")
///__________________________________________________

panFilesList=listfiles(folder(""),"????KASX")
activeFiles=info("files")
activeWindows=info("windows")
allvariables="Global variables"+¶+¶+info("globalvariables")+¶+¶+"local variables"+¶+¶+info("localvariables")+¶+¶+"fileglobal variables"+¶+¶+info("filevariables")+¶+¶+"window variables"+¶+¶+info("windowvariables")

//____bugcheck_______
;displaydata procNames
;displaydata procDBs
//___________________

lineNum=1
procCount=arraysize(procNames,¶)
procTextArray=""

//_______build an array of the procedure text of all the last used procedures_____
/*
Notes: this kept breaking when I tried to use arrayfilters or arraybuilds, and apparently there's 
known issues with using local variables that throws an exception about the call() procedure
because its using a subroutine using EXECTUTE to do arrayfilters 

That being said, it kept breaking even after turning the variables global, so now it's a loop
*/
//________________________________________________________________________________

loop
    vDb=array(procDBs,lineNum,¶)
    vProc=array(procNames,lineNum,¶)
        getproceduretext vDb,vProc,procText
        procTextArray=vProc+¶+¶+procText+¶+procTextArray  //format: Name of Procedure, two returns, text from the proc, then the last thing added put on the end
    lineNum=lineNum+1
while lineNum<procCount


//_________________Make code into word array______________________//
/*
this function makes two arrays similar enough to compare to find out
which of the active variables was in the procedures that were recently called
gets rid of the most common characters in the text and replaces them with ; to give the other functions
a separator to work with

This was done because there's like 30 variables that are only Panorama's that also gets included in the INFO("xVARIABLE")
calls, and those aren't really useful for bugfixing procedures. 

//______________get out extra characters, but retrain spaces between words using a semicolon__________
note: there were also pipes '||' with curley brackets between them, but those break the SETPROCEDURETEXT statement, so I had to take them out of the "GetMacros" version
*/
strippedText=replacemultiple(procTextArray,
“.||?||!||,||;||:||-||_||(||)||[||]||"||'||+||¶||¬||/||=||*||" "|| ||”,
“;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||;||”,
"||")

strippedText=stripchar(strippedText,"AZaz09;")
arraystrip strippedText,";"

//_____Change the format of the array into a ¶ one_______
strippedText=replace(strippedText,";",¶)
arraydeduplicate strippedText,strippedText,¶

//________get variablelist into a cleaner version_____
usedvariables=arraystrip(allvariables,¶)

//__________do a comparison for whats in both of them and put that in printVariables
arrayboth strippedText, usedvariables, ¶, printVariables

//_______Print Check_____
;displaydata printVariables


////____________Error Log____________________
/*
The short form of what gets displayed to the user specifically to be added to the 
errorDictionary. You can get this full log by calling 
DISPLAYDATA errorDictionary
*/
DictNameToday=superdatepattern(supernow(),"mm/dd/yy@", "hh:mm" )
setdictionaryvalue errorDictionary,DictNameToday, 
"Error: '"+mostRecentProc+"' created an error."
+¶+¶+
"ErrorCode: "+getError
+¶+¶+¶+
"folder in use: "
+¶+
folderArray
+¶+¶+¶+
"current Pan files in that folder are: "
+¶+
panFilesList
+¶+¶+¶+
"currently open files are: "
+¶+
activeFiles
+¶+¶+¶+
"currently open windows are: "
+¶+
activeWindows
+¶+¶+¶+
"last procedures run were"
+¶+
procList
+¶+¶+¶+
"text of non-design/form procedures:"
+¶+
procTextArray
+¶+¶+¶+
"variables used in last macros:"
+¶+
printVariables

///__________Future feature_____________
/*
Give the user instructions on what to do based on the error
*/
errorList="array of errors to give advice about"
//_______________________________________


////_____________ErrorDisplay for user________________________________________

displaydata "Error: '"+mostRecentProc+"' procedure/macro created an error."
+¶+¶+
"ErrorCode: "+getError
+¶+¶+
"Warning! If you click OK the macro will continue without fixing
the error. Proceed with caution, or click Stop instead."
+¶+¶+
"Click 'stop' to end the macro here and try what you were doing again"
+¶+¶+
"If the problem persists, use the 'COPY' button, paste this error in an e-mail 
and send it to: tech-support@fedcoseeds.com with a description of what happened



_______________________________________________________________________________"
+¶+¶+¶+
"---------------------------------------------------
THE FOLLOWING LINES ARE TO HELP WITH ERROR CHECKING
---------------------------------------------------"
+¶+¶+¶+
"folder in use: "
+¶+
folderArray
+¶+¶+¶+
"current Pan files in that folder are: "
+¶+
panFilesList
+¶+¶+¶+
"currently open files are: "
+¶+
activeFiles
+¶+¶+¶+
"currently open windows are: "
+¶+
activeWindows
+¶+¶+¶+
"last procedures run were"
+¶+
procList
+¶+¶+¶+
"text of non-design/form procedures:"
+¶+
procTextArray
+¶+¶+¶+
"variables used in last macros:"
+¶+
printVariables, 
“title="Error Capture Bot 3.0" 
captionwidth=900 
size=17 
height=500 
width=1000”



/*
//_________What this error looks like___________

Error: '.GetErrorLog' procedure/macro created an error.

ErrorCode: Unknown

Warning! If you click OK the macro will continue without fixing
the error. Proceed with caution, or click Stop instead.

Click 'stop' to end the macro here and try what you were doing again

If the problem persists, use the 'COPY' button, paste this error in an e-mail 
and send it to: tech-support@fedcoseeds.com with a description of what happened



_______________________________________________________________________________


---------------------------------------------------
THE FOLLOWING LINES ARE TO HELP WITH ERROR CHECKING
---------------------------------------------------


folder in use: 
LunarWindflower:Applications:Panorama:Panorama.app:Contents:MacOS:


current Pan files in that folder are: 
ProVUE Registration.pan


currently open files are: 
Untitled


currently open windows are: 
Untitled
Untitled:.GetErrorLog


last procedures run were
.GetErrorLog	Untitled		0


text of non-design/form procedures:
.GetErrorLog

The Text of the .GetErrorLog macro would be here, but I don't wanna double this file's
Length


variables used in last macros:
activeFiles
activeWindows
allvariables
DictNameToday
errorDictionary
errorList
errorMsg
fileNeeded
folderArray
getError
lineNum
mostRecentProc
panFilesList
printVariables
procCount
procDBs
procList
procNames
procText
procTextArray
sizeCheck
smallFolderArray
strippedText
usedvariables
vDb
vProc

*/

___ ENDPROCEDURE .GetErrorLog __________________________________________________

___ PROCEDURE SeeErrorLog ______________________________________________________

    displaydata errorDictionary

___ ENDPROCEDURE SeeErrorLog ___________________________________________________

___ PROCEDURE .WaitXSeconds ____________________________________________________
local start, end,secondsToWait

secondsToWait=5
start=now()
end=start+secondsToWait
loop
    nop
while now()≤end

//_____test timer____
;message end - start

___ ENDPROCEDURE .WaitXSeconds _________________________________________________

___ PROCEDURE GetWindowSize ____________________________________________________
global newrec, rectangle1,RecTop,RecLeft,RecHeight,RecWidth,whichWin,winList2

winList2=info("windows")
superchoicedialog winList2,whichWin,“caption="Which Window do you want the size of?"”
window (whichWin)
rectangle1=info("windowrectangle")
RecTop=rtop(rectangle1)
RecLeft=rleft(rectangle1)
RecHeight=rheight(rectangle1)
RecWidth=rwidth(rectangle1)

newrec=str(RecTop)+","+str(RecLeft)+","+str(RecHeight)+","+str(RecWidth)
message "You now have the Top, Left, Height, and Width of the window. You can use the setwindow command with these numbers"
clipboard()=newrec
//top,left,height,width


___ ENDPROCEDURE GetWindowSize _________________________________________________
