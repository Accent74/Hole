аЯрЁБс                >  ўџ	                               ўџџџ       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџC e l l ( R o w N o ,   3 ) . V a l u e  
 	 	 	 	 S h e e t 1 . C e l l ( A g R o w ,   7 ) . V a l u e   =   S h e e t 1 . C e l l ( A g R o w ,   7 ) . V a l u e   +   S h e e t 1 . C e l l ( R o w N o ,   7 ) . V a l u e  
 	 	 	 	 S h e e t 1 . C e l l ( A g R o w ,   8 ) . V a l u e   =   S h e e t 1 . C e l l ( A g R o w ,   8 ) . V a l u e   +   S h e e t 1 . C e l l ( R o w N o ,   8 ) . V a l u e  
  
 	 	 	 E n d   I f  
  
 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   1 ) . C e l l P a r a m   =   a D a t a ( 4 ,   i )  
  
 	 	 	 N D a y s   =   ( a D a t a ( 8 ,   i )   -   a D a t a ( 5 ,   i ) )  
 	  
 	 	 	 I f   N D a y s   >   0   T h e n  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   9 ) . V a l u e   =   " ' "   &   N D a y s  
 	 	 	 E n d   I f  
  
 	 	 	 N D a y s   =   ( a D a t a ( 8 ,   i )   -   d E n d )  
  
 	 	 	 I f   N D a y s   <   0   T h e n  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   1 0 ) . V a l u e   =   " ' "   &   A b s ( N D a y s )  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   1 0 ) . F o r e C o l o r   =   v b R e d  
 	 	 	 E n d   I f  
  
 	 	 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   3 ) . V a l u e   =   S h e e t 1 . C e l l ( S T A R T _ R O W ,   3 ) . V a l u e   +   S h e e t 1 . C e l l ( R o w N o ,   3 ) . V a l u e  
 	 	 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   7 ) . V a l u e   =   S h e e t 1 . C e l l ( S T A R T _ R O W ,   7 ) . V a l u e   +   S h e e t 1 . C e l l ( R o w N o ,   7 ) . V a l u e  
 	 	 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   8 ) . V a l u e   =   S h e e t 1 . C e l l ( S T A R T _ R O W ,   8 ) . V a l u e   +   S h e e t 1 . C e l l ( R o w N o ,   8 ) . V a l u e  
  
 	 	 	 W i t h   S h e e t 1 . C e l l ( R o w N o ,   5 )  
 	 	 	 	 . H L i n k   =   T r u e  
 	 	 	 	 . C e l l P a r a m   =   a D a t a ( 4 ,   i )  
 	 	 	 E n d   W i t h  
  
 	 	 N e x t  
 	 E n d   I f  
  
 	 s h e e t 1 . r a n g e ( 1 ,   S T A R T _ R O W   -   2 ,   S h e e t 1 . C o l u m n s ,   S h e e t 1 . R o w s ) . A u t o F i t   1   +   2  
 	 S h e e t 1 . R a n g e ( 3 ,   S T A R T _ R O W   +   1 ,   3 ,   S h e e t 1 . R o w s ) . A l i g n m e n t   a c R i g h t  
 	 S h e e t 1 . R a n g e ( 7 ,   S T A R T _ R O W   +   1 ,   S h e e t 1 . C o l u m n s ,   S h e e t 1 . R o w s ) . A l i g n m e n t   a c R i g h t  
  
 	 W i t h   S h e e t 1 . R a n g e ( 1 ,   S T A R T _ R O W   +   1 ,   S h e e t 1 . C o l u m n s ,   S h e e t 1 . R o w s )  
 	 	 . S t r i p e    
 	 	 . S e t B o r d e r   a c B r d G r i d , , R G B ( 1 2 5 ,   1 5 8 ,   1 9 2 )  
 	 E n d   W i t h  
  
 E n d   S u b  
  
 S u b   S h e e t 1 _ O n C e l l N a v i g a t e ( R o w ,   C o l u m n )  
 	 D i m   D o c I D  
  
 	 D o c I D   =   S h e e t 1 . C e l l ( R o w ,   C o l u m n ) . C e l l P a r a m  
  
 	 I f   D o c I D   < >   0   T h e n   	 w o r k a r e a . O p e n D o c u m e n t   D o c I D  
 E n d   S u b <     ш                Ь Arial 1Ђ'Hz   ф№    hё-єьrџџџџ      џџџџ                 џџ 
 CSheetPageџўџS h e e t 1 џўџ8AB  1 
         61    7                                    
 (   V   Ѓ   $   +   K   G   Q   P   \                            џџ  CRow џџ  CCell6џўџ џўџ     џџ 2   515B>@A:0O  704>;65==>ABL                  џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                      џўџ џўџ     џџ      =0  40BC                  џўџ џўџ     џџ $   0 1   45:01@O  2 0 1 7   3.                   џўџ џўџ     џџ                     џўџ џўџ     џџ                     џўџ џўџ     џџ                     џўџ џўџ     џџ                     џўџ џўџ     џџ                     џўџ џўџ     џџ                     
 џўџ џўџ     џџ      ?>  AG5BC                  џўџ џўџ     џџ    6 3 2 3                      џўџ џўџ     џџ                     џўџ џўџ     џџ                     џўџ џўџ     џџ                     џўџ џўџ     џџ                     џўџ џўџ     џџ                     џўџ џўџ     џџ                       џўџ џўџ " џџџџ      E U R                   
 џўџ џўџ % џџ     >:C?0B5;L                  џўџ џўџ % џџ                         џўџ џўџ % џџ  2   1I0O  AC<<0  704>;65==>AB8                  џўџ џўџ % џџ  2    0AH8D@>2:0  704>;65==>AB8                  џўџ џўџ % џџ                      џўџ џўџ % џџ                      џўџ џўџ % џџ                      џўџ џўџ % џџ                      џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     
 џўџ џўџ % џџ                         џўџ џўџ % џџ       08<5=>20=85                  џўџ џўџ % џџ                        џўџ џўџ % џџ       0B0                  џўџ џўџ % џџ    
   ><5@                  џўџ џўџ % џџ       ?;0B8BL  4>                  џўџ џўџ % џџ       @>A@>G5=>                  џўџ џўџ % џџ       45B  >?;0BK                  џўџ џўџ  џџ       BA@>G:0,   4=.                   џўџ џўџ  џџ       @>A@>G5=>,   4=.                   
 џўџ џўџ $        B>3>:                   џўџ џўџ %                            џўџ џўџ &                           џўџ џўџ %                             џўџ џўџ %                             џўџ џўџ %                             џўџ џўџ &                           џўџ џўџ &                           џўџ џўџ   џџ                       џўџ џўџ   џџ                       d   d   d   d   џўџ <     ш                Ь Arial /аr їл0  л   Аль№џўџ <     ш                Ь Arial        xНэit їлл   џўџ <     ш                Ь Arial а[   +Ћ w  "    И~еџўџ <     ш                Ь Arial /аr їл0  л   Аль№џўџ <     ш                Ь Arial        xНэit їлл   џўџ <     ш                Ь Arial а[   +Ћ w  "    И~е    2   2       <     ш            џ   Ь Arial ь№ь№    ШївlПVаrАл  <     ш               Ь Arial ь№ь№    ШївlПVаrАл	    џџр    ddd    ddd    ddd    ddd  џџџ    }Р    }Р    }Р    }Р  рџџ    }Р    }Р    }Р    }Р   џџџџ   }Р    }Р    }Р    }Р  ццњ    }Р    }Р    }Р    }Р  џџр                                  џџр    }Р    }Р    }Р    }Р  рџџ                                  џџџ                                               џўџ џўџ " џџ                       џўџ џўџ "           KР                  d   d   d   d   џўџ <     ш                Ь Arial /аrф
мT  л   АлШіџўџ <     ш                Ь Arial        TУэitф
мл   џўџ <     ш                Ь Arial MЗOxа[   +Ћ w  "    џўџ <     ш                Ь Arial /аrф
мT  л   АлШіџўџ <     ш                Ь Arial        TУэitф
мл   џўџ <     ш                Ь Arial MЗOxа[   +Ћ w  "        2   2       <     ш            џ   Ь Arial ШіШі    ШївHХVаrАл  <     ш               Ь Arial ШіШі    ШївHХVаrАл	    џџр    ddd    ddd    ddd    ddd  џџџ    }Р    }Р    }Р    }Р  рџџ    }Р    }Р    }Р    }Р   џџџџ   }Р    }Р    }Р    }Р  ццњ    }Р    }Р    }Р    }Р  џџр                                  џџр    }Р    }Р    }Р    }Р  рџџ                                  џџџ                                   џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџR o o t   E n t r y                                               џџџџџџџџ                               рйн^cг   Р      C o n t e n t s                                                  џџџџ   џџџџ                                       H+       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                                                                                                  џџџџџџџџџџџџ                                                R o o t   E n t r y                                               џџџџџџџџ                               pOЌXcг         C o n t e n t s                                                  џџџџ   џџџџ                                       $/       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        e                                                                          џџџџџџџџџџџџ                                                                        	   
                  ўџџџџџџџўџџџ§џџџўџџџўџџџ            #   џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ$   %       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   5        x                   Ќ      И      Ф      а   	   м   
   ш      є                       (     у        ST74_DBReportCurs.ash                                                           4   @   @ѓОЎ   @   #ЅXcг@   РЖІ'!Бв      Ръіхэђ 7.4                                                                                                                                                                               џўџџшO p t i o n   E x p l i c i t  
 ' # i n c l u d e   " H L 7 4 _ A D O . a v b "  
 ' # i n c l u d e   " H L 7 4 _ C o n s t . a v b "  
 ' # i n c l u d e   " H L 7 4 _ C o m m o n . a v b "  
  
 C o n s t   S T A R T _ R O W   =   7  
 C o n s t   M A X _ C O L U M N S   =   1 0  
  
    
 S u b   S h t B o o k _ O n L o a d  
 	 S h e e t 1 . R o w s   =   S T A R T _ R O W  
 	 S h e e t 1 . C o l u m n s   =   M A X _ C O L U M N S  
 	 S h e e t 1 . C e l l ( 4 ,   2 ) . V a l u e   =   " "  
  
 	 S h e e t 1 . D i s p l a y H e a d i n g s   =   F a l s e  
 	 S h e e t 1 . D i s p l a y G r i d   =   F a l s e  
  
 	 W i t h   w o r k a r e a . s i t e  
 	 	 I f   . k i n d   =   a c A c c o u n t   T h e n  
 	 	 	 M a k e R e p o r t   . I D  
 	 	 E n d   I f  
 	 E n d   W i t h  
 E n d   S u b  
  
 S u b   M a k e R e p o r t ( A c c I D )  
 	 D i m   a D a t a ,   a P r m ,   i ,   R o w N o ,   C u r I D  
 	 D i m   A g I D ,   d E n d ,   A g R o w ,   N D a y s ,   A c c  
  
 	 S h e e t 1 . R o w s   =   S T A R T _ R O W  
 	 S h e e t 1 . C o l u m n s   =   M A X _ C O L U M N S  
  
 	 S h e e t 1 . C e l l ( 3 ,   2 ) . V a l u e   =   f o r m a t d a t e 2 ( w o r k a r e a . p e r i o d . e n d ,   " d d   m m m m   y y y y   3. " )  
  
 	 S e t   A c c   =   w o r k a r e a . a c c o u n t ( A c c I D )  
 	 S h e e t 1 . C e l l ( 4 ,   2 ) . V a l u e   =   A c c . D r a w T e x t ( 3 ) 	  
  
 	 C u r I D   =   c o m _ g e t p a r a m v a l u e ( A c c ,   p r m A c c C u r r e n c y ,   1 )  
  
 	 S h e e t 1 . C e l l ( S T A R T _ R O W   -   3 ,   M A X _ C O L U M N S ) . V a l u e   =   i i f ( C u r I D   < >   1 ,   w o r k a r e a . C u r ( C u r I D ) . S h o r t N a m e ,   " " )  
 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   3 ) . V a l u e   =   0 	  
 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   7 ) . V a l u e   =   0  
 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   8 ) . V a l u e   =   0  
  
 	 d E n d   =   W o r k a r e a . P e r i o d . E n d  
  
 	 c a p t i o n   =   S h e e t 1 . C e l l ( 2 ,   1 ) . V a l u e   &   "   =0  "   &   S h e e t 1 . C e l l ( 3 ,   2 ) . V a l u e   &   "   ?>  AG5BC  "   &   S h e e t 1 . C e l l ( 4 ,   2 ) . V a l u e  
  
 	 a P r m   =   A r r a y ( A c c I D ,   C u r I D ,   d E n d ,   W o r k a r e a . M y C o m p a n y . I D )  
  
 	 I f   Q u e r y ( " S T 7 4 _ D B R e p o r t C u r s " ,   a P r m ,   a D a t a )   T h e n  
 	 	 A                         	   
            ўџџџџџџџџџџџўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ§џџџўџџџўџџџ          !   "   &   џџџџџџџџџџџџ'       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ                  ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   m        x            и      ф      №      ќ        	     
         ,     8     D    X      `     у     P   Фхсхђю№ёърџ\ъ№хфшђю№ёърџ чрфюыцхээюёђќ эр 01 фхърс№џ 2017 у. яю ёїхђѓ 6323 ХТаЮ                                                         5   @   АРФЎ   @   №<д^cг@   РЖІ'!Бв      Ръіхэђ 7.4                                                                                                                       џўџџшO p t i o n   E x p l i c i t  
 ' # i n c l u d e   " H L 7 4 _ A D O . a v b "  
 ' # i n c l u d e   " H L 7 4 _ C o n s t . a v b "  
 ' # i n c l u d e   " H L 7 4 _ C o m m o n . a v b "  
  
 C o n s t   S T A R T _ R O W   =   7  
 C o n s t   M A X _ C O L U M N S   =   1 0  
  
    
 S u b   S h t B o o k _ O n L o a d  
 	 S h e e t 1 . R o w s   =   S T A R T _ R O W  
 	 S h e e t 1 . C o l u m n s   =   M A X _ C O L U M N S  
 	 S h e e t 1 . C e l l ( 4 ,   2 ) . V a l u e   =   " "  
  
 	 S h e e t 1 . D i s p l a y H e a d i n g s   =   F a l s e  
 	 S h e e t 1 . D i s p l a y G r i d   =   F a l s e  
  
 	 W i t h   w o r k a r e a . s i t e  
 	 	 I f   . k i n d   =   a c A c c o u n t   T h e n  
 	 	 	 M a k e R e p o r t   . I D  
 	 	 E n d   I f  
 	 E n d   W i t h  
 E n d   S u b  
  
 S u b   M a k e R e p o r t ( A c c I D )  
 	 D i m   a D a t a ,   a P r m ,   i ,   R o w N o ,   C u r I D  
 	 D i m   A g I D ,   d E n d ,   A g R o w ,   N D a y s ,   A c c  
  
 	 S h e e t 1 . R o w s   =   S T A R T _ R O W  
 	 S h e e t 1 . C o l u m n s   =   M A X _ C O L U M N S  
  
 	 S h e e t 1 . C e l l ( 3 ,   2 ) . V a l u e   =   f o r m a t d a t e 2 ( w o r k a r e a . p e r i o d . e n d ,   " d d   m m m m   y y y y   3. " )  
  
 	 S e t   A c c   =   w o r k a r e a . a c c o u n t ( A c c I D )  
 	 S h e e t 1 . C e l l ( 4 ,   2 ) . V a l u e   =   A c c . D r a w T e x t ( 3 ) 	  
  
 	 C u r I D   =   c o m _ g e t p a r a m v a l u e ( A c c ,   p r m A c c C u r r e n c y ,   1 )  
  
 	 S h e e t 1 . C e l l ( S T A R T _ R O W   -   3 ,   M A X _ C O L U M N S ) . V a l u e   =   i i f ( C u r I D   < >   1 ,   w o r k a r e a . C u r ( C u r I D ) . S h o r t N a m e ,   " " )  
 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   3 ) . V a l u e   =   0 	  
 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   7 ) . V a l u e   =   0  
 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   8 ) . V a l u e   =   0  
  
 	 d E n d   =   W o r k a r e a . P e r i o d . E n d  
  
 	 c a p t i o n   =   S h e e t 1 . C e l l ( 2 ,   1 ) . V a l u e   &   "   =0  "   &   S h e e t 1 . C e l l ( 3 ,   2 ) . V a l u e   &   "   ?>  AG5BC  "   &   S h e e t 1 . C e l l ( 4 ,   2 ) . V a l u e  
  
 	 a P r m   =   A r r a y ( A c c I D ,   C u r I D ,   d E n d ,   W o r k a r e a . M y C o m p a n y . I D )  
  
 	 I f   Q u e r y ( " S T 7 4 _ D B R e p o r t C u r s " ,   a P r m ,   a D a t a )   T h e n  
 	 	 A g I D   =   0  
  
 	 	 F o r   i   =   0   T o   U B o u n d ( a D a t a ,   2 )  
 	 	 	 R o w N o   =   S h e e t 1 . I n s e r t R o w 2  
  
 	 	 	 I f   A g I D   < >   a D a t a ( 0 ,   i )   T h e n  
 	 	 	 	 A g R o w   =   R o w N o  
 	 	 	 	 A g I D   =   a D a t a ( 0 ,   i )  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   1 ) . V a l u e   =   a D a t a ( 1 ,   i )  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   2 ) . V a l u e   =   a D a t a ( 2 ,   i )  
 	 	 	 	 S h e e t 1 . g I D   =   0  
  
 	 	 F o r   i   =   0   T o   U B o u n d ( a D a t a ,   2 )  
 	 	 	 R o w N o   =   S h e e t 1 . I n s e r t R o w 2  
  
 	 	 	 I f   A g I D   < >   a D a t a ( 0 ,   i )   T h e n  
 	 	 	 	 A g R o w   =   R o w N o  
 	 	 	 	 A g I D   =   a D a t a ( 0 ,   i )  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   1 ) . V a l u e   =   a D a t a ( 1 ,   i )  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   2 ) . V a l u e   =   a D a t a ( 2 ,   i )  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   3 ) . V a l u e   =   a D a t a ( 3 ,   i )  
 	 	 	 	 S h e e t 1 . R a n g e ( 1 ,   R o w N o ,   S h e e t 1 . C o l u m n s ,   R o w N o ) . F o n t . B o l d   =   T r u e  
 	 	 	 E n d   I f  
  
 	 	 	 I f   a D a t a ( 4 ,   i )   < >   0   T h e n  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   4 ) . V a l u e   =   a D a t a ( 5 ,   i )  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   5 ) . V a l u e   =   a D a t a ( 6 ,   i )  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   6 ) . V a l u e   =   a D a t a ( 8 ,   i )  
  
 	 	 	 	 I f   a D a t a ( 8 ,   i )   <   d E n d   T h e n  
 	 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   7 ) . V a l u e   =   a D a t a ( 7 ,   i )  
 	 	 	 	 E l s e  
 	 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   8 ) . V a l u e   =   a D a t a ( 7 ,   i )  
 	 	 	 	 E n d   I f  
  
 	 	 	 	 S h e e t 1 . C e l l ( A g R o w ,   3 ) . V a l u e   =   S h e e t 1 . C e l l ( A g R o w ,   3 ) . V a l u e   +   S h e e t 1 . C e l l ( R o w N o ,   3 ) . V a l u e   =   a D a t a ( 3 ,   i )  
 	 	 	 	 S h e e t 1 . R a n g e ( 1 ,   R o w N o ,   S h e e t 1 . C o l u m n s ,   R o w N o ) . F o n t . B o l d   =   T r u e  
 	 	 	 E n d   I f  
  
 	 	 	 I f   a D a t a ( 4 ,   i )   < >   0   T h e n  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   4 ) . V a l u e   =   a D a t a ( 5 ,   i )  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   5 ) . V a l u e   =   a D a t a ( 6 ,   i )  
 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   6 ) . V a l u e   =   a D a t a ( 8 ,   i )  
  
 	 	 	 	 I f   a D a t a ( 8 ,   i )   <   d E n d   T h e n  
 	 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   7 ) . V a l u e   =   a D a t a ( 7 ,   i )  
 	 	 	 	 E l s e  
 	 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   8 ) . V a l u e   =   a D a t a ( 7 ,   i )  
 	 	 	 	 E n d   I f  
  
 	 	 	 	 S h e e t 1 . C e l l ( A g R o w ,   3 ) . V a l u e   =   S h e e t 1 . C e l l ( A g R o w ,   3 ) . V a l u e   +   S h e e t 1 . 