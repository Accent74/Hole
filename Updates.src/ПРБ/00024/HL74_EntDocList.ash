аЯрЁБс                >  ўџ	                               ўџџџ       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџN o ,   2 ) . M u l t i L i n e   =   T r u e  
 	 	 	 	 	 	 I f   E n t D b S u m   < >   0   T h e n   S h e e t 1 . C e l l ( R o w N o ,   4 ) . V a l u e   =   E n t D b S u m  
 	 	 	 	 	 	 I f   E n t C r S u m   < >   0   T h e n   S h e e t 1 . C e l l ( R o w N o ,   5 ) . V a l u e   =   E n t C R S u m  
  
 	 	 	 	 	 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   4 ) . V a l u e   =   S h e e t 1 . C e l l ( S T A R T _ R O W ,   4 ) . V a l u e   +   E n t D b S u m  
 	 	 	 	 	 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   5 ) . V a l u e   =   S h e e t 1 . C e l l ( S T A R T _ R O W ,   5 ) . V a l u e   +   E n t C r S u m  
  
 	 	 	 	 	 	 W i t h   S h e e t 1 . C e l l ( R o w N o ,   2 )  
 	 	 	 	 	 	 	 . C e l l P a r a m   =   O p . I D  
 	 	 	 	 	 	 	 . H L i n k   =   T r u e  
 	 	 	 	 	 	 E n d   W i t h  
 	  
 	 	 	 	 	 E n d   I f  
 	 	 	 	 E n d   I f  
 	 	 	 N e x t  
  
 	 	 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   M A X _ C O L U M N S ) . V a l u e   =   S h e e t 1 . C e l l ( S T A R T _ R O W ,   M A X _ C O L U M N S   -   3 ) . V a l u e   +   _  
 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   M A X _ C O L U M N S   -   2 ) . V a l u e   -   _  
 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   M A X _ C O L U M N S   -   1 ) . V a l u e  
  
 	 	 	 W i t h   S h e e t 1 . R a n g e ( 1 ,   S T A R T _ R O W   +   1 ,   M A X _ C O L U M N S ,   S h e e t 1 . R o w s )  
 	 	 	 	 . A u t o F i t   1  
 	 	 	 	 . S t r i p e  
 	 	 	 	 . S e t B o r d e r   a c B r d G r i d , ,   R G B ( 1 2 5 ,   1 5 8 ,   1 9 2 )  
 	 	 	 E n d   W i t h  
  
 	 	 	 S h e e t 1 . R a n g e ( 3 ,   S T A R T _ R O W   +   1 ,   M A X _ C O L U M N S ,   S h e e t 1 . R o w s ) . A l i g n m e n t   =   a c R i g h t  
  
 	 	 E n d   I f  
 	 E n d   W i t h  
 E n d   S u b  
  
 F u n c t i o n   F i n d A c c ( O p ,   d A c c ,   A c c I D ,   E n t I D ,   E n t D b S u m ,   E n t C r S u m )  
 	 D i m   i ,   j  
  
 	 E n t D b S u m   =   0  
 	 E n t C r S u m   =   0  
  
 	 F o r   i   =   1   T o   O p . T r a n s C o u n t  
 	 	 W i t h   O p . T r a n s L i s t ( i )  
 	 	 	 F o r   j   =   1   T o   . R o w s  
 	 	 	 	 W i t h   . I t e m ( j )  
 	 	 	 	 	 I f   . E n t I D   =   E n t I D   T h e n  
 	 	 	 	 	 	 I f   d A c c . E x i s t s ( . A c c D B I D )   T h e n  
 	 	 	 	 	 	 	 E n t D b S u m   =   E n t D b S u m   +   . Q t y  
 	 	 	 	 	 	 E n d   I f  
  
 	 	 	 	 	 	 I f   d A c c . E x i s t s ( . A c c C R I D )   T h e n  
 	 	 	 	 	 	 	 E n t C r S u m   =   E n t C r S u m   +   . Q t y  
 	 	 	 	 	 	 E n d   I f  
 	 	 	 	 	 E n d   I f  
 	 	 	 	 E n d   W i t h  
 	 	 	 N e x t  
 	 	 E n d   W i t h  
 	 N e x t  
 	 	 	 	 	  
 E n d   F u n c t i o n  
  
 F u n c t i o n   c o m _ C a l c E n t R e s t ( A c c I D ,   E n t I D ,   O n D a t e )  
 	 D i m   a P r m ,   a D a t a  
  
 	 a P r m   =   A r r a y ( A c c I D ,   E n t I D ,   O n D a t e ,   W o r k a r e a . M y C o m p a n y . I D )  
  
 	 I f   Q u e r y ( " H L 7 4 _ R e s t B y E n t A l l A g " ,   a P r m ,   a D a t a )   T h e n  
 	 	 c o m _ C a l c E n t R e s t   =   c h e c k n u l l ( a D a t a ( 0 ,   0 ) ,   0 )  
 	 E l s e  
 	 	 c o m _ C a l c E n t R e s t   =   0  
 	 E n d   I f  
  
 E n d   F u n c t i o n  
  
 S u b   S h e e t 1 _ O n C e l l N a v i g a t e ( R o w ,   C o l u m n )  
 	 D i m   I D  
  
 	 i d   =   S h e e t 1 . C e l l ( R o w ,   C o l u m n ) . C e l l P A r a m  
 	 I f   i d   < >   0   T h e n   W o r k A r e a . O p e n D o c u m e n t ( i d )  
 E n d   S u b  
 <     ш                Ь Arial Щewz    №    Є№-єlRjOzа      џџџџ                 џџ 
 CSheetPageџўџS h e e t 1 џўџ8AB  1    
      61    7                                     7   ]  k   /   .   d   
                             .   
 џџ  CRow џџ  CCell6џўџ џўџ   џџџџ                       џўџ џўџ     џџ F   BG5B  >  42865=88  B>20@>2  =0  A:;040E                  џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                      џўџ џўџ   џџџџ      70                  џўџ џўџ= w o r k a r e a . p e r i o d . t i t l e     џџ
                       џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                      џўџ џўџ   џџџџ      ?>  AG5BC                  џўџ џўџ     џџ (   2 8 1   ">20@K  =0  A:;045                  џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                     џўџ џўџ   џџџџ                      џўџ џўџ 5 џџ    
   !  ?/ ?                  џўџ џўџ 5 џџ    &   ">20@/ !:40/ >:C<5=B                  џўџ џўџ 5 џџ    "   AB0B>:  =0  =0G0;>                  џўџ џўџ 5 џџ       @8E>4                  џўџ џўџ 5 џџ        0AE>4                  џўџ џўџ 5 џџ        AB0B>:  =0  :>=5F                   џўџ џўџ      $   04286:0  <51.   E@><                  џўџ џўџ                           џўџ џўџ "       0W                       џўџ џўџ "                              џўџ џўџ "                              џўџ џўџ "        C!                        џўџ џўџ   џџ                       џўџ џўџ  џџ    r   2 ">20@K.    50;870F8O    !  6 3 9 1   >B  1 0 . 0 9 . 1 9 / >:C?0B5;L                 џўџ џўџ " џџ                       џўџ џўџ " џџ                       џўџ џўџ " џџ    0u                        џўџ џўџ " џџ                        џўџ џўџ   џџ	                       џўџ џўџ  џџ	    r   2 ">20@K.    50;870F8O    !  6 9 2 4   >B  1 6 . 0 9 . 1 9 / >:C?0B5;L                 џўџ џўџ " џџ	                       џўџ џўџ " џџ	                       џўџ џўџ " џџ	    @                        џўџ џўџ " џџ	                        џўџ џўџ   џџ                       џўџ џўџ  џџ    Р   2 ">20@K.   >ABC?;5=85.   >ABC?;5=85  B>20@0  >B  ?>AB02I8:0    !  "  2 7 7   >B  1 8 . 0 9 . 1 9 /   " !   !""                  џўџ џўџ " џџ                       џўџ џўџ " џџ                           џўџ џўџ " џџ                       џўџ џўџ " џџ                        џўџ џўџ   џџ	                       џўџ џўџ  џџ	    
  2 ">20@K.    50;870F8O  DC@=8BC@0  !  !  7 3 6 1     >B  1 9 . 0 9 . 1 9 / 0;5=B8=  8:B>@>28G  >@>  0 7 1   3 7 9   2 9   9 7   C;.   C918H520  6 0   " B0A"   ( !B@>945@52=O)                  џўџ џўџ " џџ	                       џўџ џўџ " џџ	                       џўџ џўџ " џџ	                            џўџ џўџ " џџ	                       d   d   d   d   џўџ <     ш                Ь Arial /PRЬУФ  @хУ   hфУ(№џўџ <     ш                Ь Arial        ДМэi:RЬУ@хУ   џўџ <     ш                Ь Arial )K;аш8тX;т             ѕuџўџ <     ш                Ь Arial /PRЬУФ  @хУ   hфУ(№џўџ <     ш                Ь Arial        ДМэi:RЬУ@хУ   џўџ <     ш                Ь Arial )K;аш8тX;т             ѕu    2   2       <     ш               Ь Arial (№(№    8hСЈОVPRhфУ  <     ш               Ь Arial (№(№    8hСЈОVPRhфУ
   	 рџџ    }Р    }Р    }Р    }Р  џџџ    }Р    }Р    }Р    }Р  џџр    }Р    }Р    }Р    }Р  ццњ    }Р    }Р    }Р    }Р  рџџ                                  џџџ                                  џџр    РРР    РРР    РРР    РРР  ццњ    РРР    РРР    РРР    РРР  ццњ                                   џџр                                                                h/       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        a                                                                          џџџџџџџџџџџџ                                                R o o t   E n t r y                                               џџџџџџџџ                               аGИ~е         C o n t e n t s                                                  џџџџ   џџџџ                                       к0       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        a                                                                          џџџџџџџџџџџџ                                                R o o t   E n t r y                                               џџџџџџџџ                               h~е         C o n t e n t s                                                  џџџџ   џџџџ                                       Т0       S u m m a r y I n f o r m a t i o n                           (  џџџџџџџџџџџџ                                        a                                                                          џџџџџџџџџџџџ                                                                        	   
                     ўџџџџџџџўџџџ§џџџўџџџўџџџ          џџџџџџџџџџџџџџџџџџџџџџџџџџџџ!   "   #   &   џџџџџџџџ    џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   1        x                  Ј      Д      Р      Ь   	   и   
   ф      №      ќ                $     у        HL74_EntDocList.ash                                                         4   @   4^Ы  @    |h~е@   х!Їе      Ръіхэђ 7.4                                                   џўџO p t i o n   E x p l i c i t  
 <     ш                Ь Arial СаШы    ь    ь9"охџџџџ      џџџџ                 џўџџ[' # i n c l u d e   " H L 7 4 _ A D O . a v b "  
 ' # i n c l u d e   " H L 7 4 _ C o m m o n . a v b "  
  
 O p t i o n   E x p l i c i t  
  
 C o n s t   S T A R T _ R O W   =   6  
 C o n s t   M A X _ C O L U M N S   =   6  
  
 D i m   A c c I D ,   E n t I D  
 D i m   d A c c  
  
 S u b   S h t B o o k _ O n L o a d  
 	 W i t h   W o r k a r e a . S i t e  
 	 	 I f   . K i n d   =   a c E n t i t y   T h e n  
 	 	 	 S e t   E n t   =   W o r k a r e a . E n t i t y ( . I D )  
 	 	 	 A c c I D   =   c o m _ G e t L e g a c y A c c I D ( E n t )  
 	 	 E n d   I f  
 	 E n d   W i t h  
  
 	 I f   A c c I D   =   0   T h e n   A c c I D   =   W o r k a r e a . G e t A c c I D ( " 2 8 " )  
 	 S e t   d A c c   =   L o a d S u b A c c ( A c c I D )  
 	  
 	 M a k e R e p o r t   A c c I D ,   d A c c  
  
 E n d   S u b  
  
  
 F u n c t i o n   L o a d S u b A c c ( A c c I D )  
 	 D i m   d i c t ,   i  
  
 	 S e t   d i c t   =   C r e a t e L i b O B j e c t ( " M a p " )  
  
 	 I f   A c c I D   < >   0   T h e n  
 	 	 S h e e t 1 . C e l l ( 4 ,   2 ) . V a l u e   =   W o r k a r e a . A c c o u n t ( A c c I D ) . D r a w T e x t ( 3 )  
 	 	 D i c t ( A c c I D )   =   A c c I D  
  
 	 	 W i t h   w o r k a r e a . A c c o u n t ( A c c I D ) . N e s t e d  
 	 	 	 F o r   i   =   1   T o   . C o u n t  
 	 	 	 	 W i t h   . I t e m ( i )  
 	 	 	 	 	 I f   N o t   D i c t . E x i s t s ( . I D )   T h e n   D i c t ( . I D )   =   . I D                         	   
                     ўџџџўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ§џџџўџџџўџџџ         $   џџџџџџџџџџџџџџџџ%   '   џџџџ(       џџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџ               ўџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџџўџ   рђљOhЋ +'Гй   рђљOhЋ +'Гй0   1        x                  Ј      Д      Р      Ь   	   и   
   ф      №      ќ                $     у        HL74_EntDocList.ash                                                         5   @   рNЎiЫ  @   @wИ~е@   х!Їе      Ръіхэђ 7.4                                                   џўџO p t i o n   E x p l i c i t  
 <     ш                Ь Arial СаШы    ь    ь9"охџџџџ      џџџџ                 џўџџg' # i n c l u d e   " H L 7 4 _ A D O . a v b "  
 ' # i n c l u d e   " H L 7 4 _ C o m m o n . a v b "  
  
 O p t i o n   E x p l i c i t  
  
 C o n s t   S T A R T _ R O W   =   6  
 C o n s t   M A X _ C O L U M N S   =   6  
  
 D i m   A c c I D ,   E n t I D  
 D i m   d A c c  
  
 S u b   S h t B o o k _ O n L o a d  
 	 D i m   E n t  
  
 	 W i t h   W o r k a r e a . S i t e  
 	 	 I f   . K i n d   =   a c E n t i t y   T h e n  
 	 	 	 S e t   E n t   =   W o r k a r e a . E n t i t y ( . I D )  
 	 	 	 A c c I D   =   c o m _ G e t L e g a c y A c c I D ( E n t )  
 	 	 E n d   I f  
 	 E n d   W i t h  
  
 	 I f   A c c I D   =   0   T h e n   A c c I D   =   W o r k a r e a . G e t A c c I D ( " 2 8 " )  
 	 S e t   d A c c   =   L o a d S u b A c c ( A c c I D )  
 	  
 	 M a k e R e p o r t   A c c I D ,   d A c c  
  
 E n d   S u b  
  
  
 F u n c t i o n   L o a d S u b A c c ( A c c I D )  
 	 D i m   d i c t ,   i  
  
 	 S e t   d i c t   =   C r e a t e L i b O B j e c t ( " M a p " )  
  
 	 I f   A c c I D   < >   0   T h e n  
 	 	 S h e e t 1 . C e l l ( 4 ,   2 ) . V a l u e   =   W o r k a r e a . A c c o u n t ( A c c I D ) . D r a w T e x t ( 3 )  
 	 	 D i c t ( A c c I D )   =   A c c I D  
  
 	 	 W i t h   w o r k a r e a . A c c o u n t ( A c c I D ) . N e s t e d  
 	 	 	 F o r   i   =   1   T o   . C o u n t  
 	 	 	 	 W i t h   . I t e m ( i )  
 	 	 	 	 	 I f   N o t   D i c t . E x i s t s ( . I D )   T h e n   D i c t ( . I D )   =   . I D  
 	 	 	 	 E n d   W i t h  
 	 	 	 N e x t  
 	 	 E n d   W i t h  
 	 E n d   I f  
  
 	 S e t   L o a d S u b A c c   =   d i c t  
 E n d   F u n c t i o n  
  
 S u b   M a k e R e p o r t ( A c c I D ,   d A c c )  
 	 D i m   d l ,   R o w N o ,   i  
 	 D i m   E n t D b S u m ,   E n t C r S u m ,   O p ,   D o c N a m e  
  
 	 S h e e t 1 . R o w s   =   S T A R T _ R O W  
 	 S h e e t 1 . C o l u m n s   =   M A X _ C O L U M N S  
 	 S h e e t 1 . C e l  
 	 	 	 	 E n d   W i t h  
 	 	 	 N e x t  
 	 	 E n d   W i t h  
 	 E n d   I f  
  
 	 S e t   L o a d S u b A c c   =   d i c t  
 E n d   F u n c t i o n  
  
 S u b   M a k e R e p o r t ( A c c I D ,   d A c c )  
 	 D i m   d l ,   R o w N o ,   i  
 	 D i m   E n t D b S u m ,   E n t C r S u m ,   O p ,   D o c N a m e  
  
 	 S h e e t 1 . R o w s   =   S T A R T _ R O W  
 	 S h e e t 1 . C o l u m n s   =   M A X _ C O L U M N S  
 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   4 ) . V a l u e   =   0  
 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   5 ) . V a l u e   =   0  
  
 	 W i t h   W o r k a r e a . S i t e  
  
 	 	 I f   . K i n d   < >   a c E n t i t y   T h e n   E x i t   S u b  
  
 	 	 E n t I D   =   . I D  
  
 	 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   1 ) . V a l u e   =   W o r k a r e a . E n t i t y ( E n t I D ) . N a m e  
 	 	 S e t   d l   =   W o r k a r e a . C r e a t e R e p o r t ( " R e p D o c L i s t " )  
  
 	 	 d l . K i n d   =   a c E n t i t y  
 	 	 d l . K i n d I D   =   E n t I D  
  
 	 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   3 ) . V a l u e   =   c o m _ C a l c E n t R e s t ( A c c I D ,   E n t I D ,   W o r k a r e a . P e r i o d . S t a r t )  
  
 	 	 I f   d l . B u i l d   T h e n  
 	 	 	 F o r   i   =   1   T o   d l . C o u n t 	  
 	 	 	 	 S e t   O p   =   w o r k a r e a . O p e r a t i o n ( d l . I t e m ( i ) . I D )  
  
 	 	 	 	 I f   O p . D o n e   T h e n  
 	 	 	 	 	 F i n d A c c   O p ,   d A c c ,   A c c I D ,   E n t I D ,   E n t D b S u m ,   E n t C r S u m  
 	  
 	 	 	 	 	 I f   ( E n t D b S u m   < >   0   O r   E n t C r S u m   < >   0 )   A n d   ( E n t D b S u m   < >   E n t C r S u m )   T h e n    
 	 	 	 	 	 	 R o w N o   =   S h e e t 1 . I n s e r t R o w 2  
 	 	 	 	 	 	 D o c N a m e   =   O p . N a m e   &   "   !  "   &   O p . D o c N o   &   F o r m a t D a t e 2 ( O p . D a t e ,   "   >B  d d . m m . y y / " )  
 	  
 	 	 	 	 	 	 I f l ( S T A R T _ R O W ,   4 ) . V a l u e   =   0  
 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   5 ) . V a l u e   =   0  
  
 	 W i t h   W o r k a r e a . S i t e  
  
 	 	 I f   . K i n d   < >   a c E n t i t y   T h e n   E x i t   S u b  
  
 	 	 E n t I D   =   . I D  
  
 	 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   1 ) . V a l u e   =   W o r k a r e a . E n t i t y ( E n t I D ) . N a m e  
 	 	 S e t   d l   =   W o r k a r e a . C r e a t e R e p o r t ( " R e p D o c L i s t " )  
  
 	 	 d l . K i n d   =   a c E n t i t y  
 	 	 d l . K i n d I D   =   E n t I D  
  
 	 	 S h e e t 1 . C e l l ( S T A R T _ R O W ,   3 ) . V a l u e   =   c o m _ C a l c E n t R e s t ( A c c I D ,   E n t I D ,   W o r k a r e a . P e r i o d . S t a r t )  
  
 	 	 I f   d l . B u i l d   T h e n  
 	 	 	 F o r   i   =   1   T o   d l . C o u n t 	  
 	 	 	 	 S e t   O p   =   w o r k a r e a . O p e r a t i o n ( d l . I t e m ( i ) . I D )  
  
 	 	 	 	 I f   O p . D o n e   T   E n t D b S u m   < >   0   A n d   E n t C R S u m   =   0   T h e n    
 	 	 	 	 	 	 	 D o c N a m e   =   D o c N a m e   &   O p . T r a n s ( 1 ) . A g F r o m B i n d . N a m e  
 	 	 	 	 	 	 E l s e  
 	 	 	 	 	 	 	 D o c N a m e   =   D o c N a m e   &   O p . T r a n s ( 1 ) . A g T o B i n d . N a m e  
 	 	 	 	 	 	 E n d   I f  
 	  
 	 	 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   2 ) . V a l u e   =   D o c N a m e  
 	 	 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   2 ) . M u l t i h e n  
 	 	 	 	 	 F i n d A c c   O p ,   d A c c ,   A c c I D ,   E n t I D ,   E n t D b S u m ,   E n t C r S u m  
 	  
 	 	 	 	 	 I f   ( E n t D b S u m   < >   0   O r   E n t C r S u m   < >   0 )   A n d   ( E n t D b S u m   < >   E n t C r S u m )   T h e n    
 	 	 	 	 	 	 R o w N o   =   S h e e t 1 . I n s e r t R o w 2  
 	 	 	 	 	 	 D o c N a m e   =   O p . N a m e   &   "   !  "   &   O p . D o c N o   &   F o r m a t D a t e 2 ( O p . D a t e ,   "   >B  d d . m m . y y / " )  
 	  
 	 	 	 	 	 	 I f   E n t D b S u m   < >   0   A n d   E n t C R S u m   =   0   T h e n    
 	 	 	 	 	 	 	 D o c N a m e   =   D o c N a m e   &   O p . T r a n s ( 1 ) . A g F r o m B i n d . N a m e  
 	 	 	 	 	 	 E l s e  
 	 	 	 	 	 	 	 D o c N a m e   =   D o c N a m e   &   O p . T r a n s ( 1 ) . A g T o B i n d . N a m e  
 	 	 	 	 	 	 E n d   I f  
 	  
 	 	 	 	 	 	 S h e e t 1 . C e l l ( R o w N o ,   2 ) . V a l u e   =   D o c N a m e  
 	 	 	 	 	 	 S h e e t 1 . C e l l ( R o w 