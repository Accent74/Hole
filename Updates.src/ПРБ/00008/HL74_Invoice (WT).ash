–ѕа°±б                >  ю€	                               ю€€€       €€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€ю€€KO p t i o n   E x p l i c i t  
 ' # i n c l u d e   " H L 7 4 _ C o n s t . a v b "  
 ' # i n c l u d e   " H L 7 4 _ F o r m s . a v b "  
 ' # i n c l u d e   " H L 7 4 _ C o m m o n . a v b "  
  
 C o n s t   G R I D _ S T A R T _ R O W   =   2 0  
 C o n s t   G R I D _ S T A R T _ C O L   =   2  
 C o n s t   M A X _ C O L U M N S   =   9  
 C o n s t   G R I D _ R O W _ B I N D   =   " R o w N o ; E n t B i n d . A r t ; E n t B i n d . N a m e ; U n i t B i n d . N a m e ; Q T Y ; P r i c e ; S u m "  
  
 D i m   O p ,   A g r e e M s c N o ,   M y C o  
  
 S u b   D e l e t e R o w s  
 	 W h i l e   S h e e t 1 . C e l l ( G R I D _ S T A R T _ R O W   +   1 ,   G R I D _ S T A R T _ C O L ) . V a l u e   < >   " "  
 	 	 S h e e t 1 . D e l e t e R o w   G R I D _ S T A R T _ R O W   +   1  
 	 W e n d  
 E n d   S u b  
  
 S u b   S e t D a t a B i n d  
 	 D i m   a D a t a B i n d ,   i ,   c o l ,   T r  
  
 	 a D a t a B i n d   =   S p l i t ( G R I D _ R O W _ B I N D ,   " ; " )  
  
 	 F o r   i   =   O p . T r a n s L i s t ( 1 ) . R o w s   T o   1   S t e p   - 1  
 	 	 S e t   T r   =   O p . T r a n s ( 1 ,   i )  
  
 	 	 F o r   C o l   =   0   T o   U B o u n d ( a D a t a B i n d )  
 	 	 	 S h e e t 1 . C e l l ( G R I D _ S T A R T _ R O W ,   C o l   +   G R I D _ S T A R T _ C O L ) . V a l u e   =   E v a l ( " t r . "   &   a D a t a B i n d ( C o l ) )  
 	 	 N e x t  
  
 	 	 I f   T r . R o w N o 2   < >   1   T h e n    
 	 	 	 S h e e t 1 . I n s e r t R o w 2 ( G R I D _ S T A R T _ R O W )  
 	 	 	 C o p y A t t r i b u t e s   G R I D _ S T A R T _ R O W ,   G R I D _ S T A R T _ R O W   +   1 ,   G R I D _ S T A R T _ C O L ,   G R I D _ S T A R T _ C O L   +   U B o u n d ( a D a t a B i n d )  
 	 	 	 S h e e t 1 . R a n g e ( G R I D _ S T A R T _ C O L ,   G R I D _ S T A R T _ R O W ,   G R I D _ S T A R T _ C O L   +   U B o u n d ( a D a t a B i n d ) ,   G R I D _ S T A R T _ R O W ) . S e t B o r d e r   a c B r d G r i d ,   2  
 	 	 E n d   I f  
 	 N e x t  
  
 E n d   S u b  
  
 S u b   C o p y A t t r i b u t e s ( R o w T o ,   R o w F r o m ,   n L e f t ,   n R i g h t )  
 	 D i m   i  
  
 	 F o r   i   =   n L e f t   T o   n R i g h t  
 	 	 S h e e t 1 . C e l l ( R o w T o ,   i ) . a l i g n m e n t   =   S h e e t 1 . C e l l ( R o w F r o m ,   i ) . a l i g n m e n t  
 	 	 S h e e t 1 . C e l l ( R o w T o ,   i ) . M u l t i l i n e   =   S h e e t 1 . C e l l ( R o w F r o m ,   i ) . M u l t i l i n e  
 	 N e x t  
  
 E n d   S u b  
  
 S u b   S h t B o o k _ O n L o a d  
 	 S e t   O p   =   S t a r t P a r a m  
 	 S e t   M y C o   =   W o r k a r e a . M y C o m p a n y  
  
 	 S h e e t 1 . C o l u m n s   =   M A X _ C O L U M N S  
 	 A g r e e M s c N o   =   c o m _ G e t M s c N o B y M s c I D ( W o r k a r e a . P a r a m s ( p r m D B A g r e e ) . V a l u e 2 )  
 	 D e l e t e R o w s  
 	 S e t D a t a B i n d  
 	 S h e e t 1 . R a n g e ( 1 ,   G R I D _ S T A R T _ R O W   -   1 	 ,   S h e e t 1 . C o l u m n s ,   S h e e t 1 . R o w s ) . A u t o F i t   1   +   2  
 	  
 E n d   S u b  
  
  
 <     и    †           ћ Arial ЇQ?az   §и     (й -фOp€€€€      €€€€                 €€ 
 CSheetPage€ю€S h e e t 1 €ю€8AB  1 	         61    7Н                                        #   o   И                                 
               
         .                                 €€  CRow   €€  CCell6€ю€ €ю€ ! €€€€Б §   =8<0=85!   >4?8A0=85  MB>9  =0:;04=>9  >7=0G05B  A>3;0A85  A  CA;>28O<8  ?>AB02:8  B>20@0.  В                 А€ю€ €ю€ ! €€€€Б    В                 А€ю€ €ю€ ! €€€€В    В                 А€ю€ €ю€ ! €€€€Г    В                 А€ю€ €ю€ ! €€€€Д    В                 А€ю€ €ю€ ! €€€€Е    В                 А€ю€ €ю€ ! €€€€Ж    В                   А А€ю€ €ю€    €€     В                 А€ю€ €ю€M= " !G5B  =0  >?;0BC  !  "   &   O p . D o c N o   &   f o r m a t d a t e 2 ( O p . D a t e ,   "   >B  d d   m m m m   y y y y   3. " )     Б   В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                 А€ю€ €ю€   €€€€Г    В                 А€ю€ €ю€   €€€€Д    В                 А€ю€ €ю€   €€€€Е    В                 А€ю€ €ю€   €€€€Ж    В                   А   А€ю€ €ю€   €€€€Б    >AB02I8::  В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€#= F o r m _ G e t A g A l t e r N a m e ( M y C o ,   O p . D a t e )    €€Б
      В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                 А€ю€ €ю€   €€€€Г    В                 А€ю€ €ю€   €€€€Д    В                 А       А€ю€ €ю€8= " 45=B8D8:0F8>==K9  :>4  N@848G5A:>3>  ;8F0:   "   &   M y C o . C o d e   €€€€Б
      В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                 А€ю€ €ю€   €€€€Г    В                 А€ю€ €ю€   €€€€Д    В                 А       А€ю€ €ю€G= " .@.   04@5A:   "   &   M y C o . C o u n t r y   &   " ,   "   &   r e p l a c e ( M y C o . A d d r e s s ,   " ; " ,   " ,   " )   €€€€Б
      В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                 А€ю€ €ю€   €€€€Г    В                 А€ю€ €ю€   €€€€Д    В                 А       А€ю€ €ю€,= " !2845B5;LAB2>  >  @538AB@0F88: "   &   M y C o . R e g N o   €€€€Б
      В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                 А€ю€ €ю€   €€€€Г    В                 А€ю€ €ю€   €€€€Д    В                 А       А€ю€ €ю€!= " @/ A  "   &   F o r m _ G e t A g A c c o u n t ( M y C o )   €€€€Б
      В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                 А€ю€ €ю€   €€€€Г    В                 А€ю€ €ю€   €€€€Д    В                   А   А€ю€ €ю€   €€€€Б    >:C?0B5;L:  В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€/= F o r m _ G e t A g A l t e r N a m e ( O p . T r a n s ( 1 ) . A g T o ,   O p . D a t e )    €€Б
      В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                 А€ю€ €ю€   €€€€Г    В                 А€ю€ €ю€   €€€€Д    В                 А       А€ю€ €ю€6= " 45=B8D8:0F8>==K9  :>4:   "   &   O p . T r a n s ( 1 ) . A g T o B i n d . C o d e   €€€€Б
      В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                 А€ю€ €ю€   €€€€Г    В                 А€ю€ €ю€   €€€€Д    В                 А       А€ю€ €ю€g= " .@.   04@5A:   "   &   O p . T r a n s ( 1 ) . A g T o B i n d . C o u n t r y   &   " ,   "   &   r e p l a c e ( O p . T r a n s ( 1 ) . A g T o B i n d . A d d r e s s ,   " ; " ,   " ,   " )   €€€€Б
      В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                 А€ю€ €ю€   €€€€Г    В                 А€ю€ €ю€   €€€€Д    В                 А       А€ю€ €ю€<= " !2845B5;LAB2>  >  @538AB@0F88: "   &   O p . T r a n s ( 1 ) . A g T o B i n d . R e g N o   €€€€Б
      В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                 А€ю€ €ю€   €€€€Г    В                 А€ю€ €ю€   €€€€Д    В                   А   А€ю€ €ю€   €€€€Б "   >3>2>@  ?>AB02:8:  В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€`= i i f ( O p . T r a n s ( 1 ) . M i s c I D ( A g r e e M s c N o )   < >   0 ,   O p . T r a n s ( 1 ) . M i s c B i n d ( A g r e e M s c N o ) . N a m e ,   " 57  4>3>2>@0" )   €€€€Б
      В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                 А€ю€ €ю€   €€€€Г    В                 А€ю€ €ю€   €€€€Д    В                   А	   А€ю€ €ю€ 5         ! В                 А€ю€ €ю€ 5         >4 В                 А€ю€ €ю€ 5      &   08<5=>20=85  B>20@0 В                 А€ю€ €ю€ 5         4.   87<.  В                 А€ю€ €ю€ 5         >;- 2> В                 А€ю€ €ю€ 5      6   &5=0,   @C1.    $          ( 157  !)  В                 А€ю€ €ю€ 5      0   !C<<0,   @C1.    $  ( 157  !)  В                 А€ю€ €ю€ 5 €€€€        В                 А   А€ю€ €ю€ % €€         В                 А€ю€ €ю€ $ €€         В                 А€ю€ €ю€  €€         В                 А€ю€ €ю€ % €€         В                 А€ю€ €ю€ & €€         В                 А€ю€ €ю€ & €€         В                 А€ю€ €ю€ & €€         В                   А       А€ю€ €ю€ " €€€€   $   A53>  =08<5=>20=89 В                 А€ю€ €ю€= O p . T r a n s L i s t ( 1 ) . R o w s "     
      В                 А€ю€ €ю€ "  €€Б .   B>3>  157  !,   @C1.    $:  В                 А€ю€ €ю€ "  €€Б    В                 А€ю€ €ю€= e v a l ( " O p . T r a n s L i s t ( 1 ) . S u m " ) "     
      В                 А       А€ю€ €ю€ " €€€€       A53>  B>20@0,   HB В                 А€ю€ €ю€= O p . T r a n s L i s t ( 1 ) . Q t y "     
      В                 А€ю€ €ю€ "  €€Б $   !  ( 0 % ) ,   @C1.    $:  В                 А€ю€ €ю€ "  €€Б    В                 А€ю€ €ю€ "           В                 А           А€ю€ €ю€ "  €€Б *   B>3>  A  !,   @C1.    $:  В                 А€ю€ €ю€ "  €€Б    В                 А€ю€ €ю€= e v a l ( " O p . S u m " ) "     
      В                   А   А€ю€ €ю€   €€€€Б    B>3>  =0  AC<<C:  В                 А€ю€ €ю€    €€Б    В                 А€ю€ €ю€C= S p e l l M o n e y ( O p . T r a n s L i s t ( 1 ) . C u r S u m ,   O p . T r a n s ( 1 ) . T r a n s C u r s ( 1 ) . C u r I D )   €€Б
      В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                 А€ю€ €ю€   €€€€Г    В                 А€ю€ €ю€   €€€€Д    В                   А   А€ю€ €ю€    €€Б    B?CAB8;:  В                 А€ю€ €ю€   €€€€Б    В                     А€ю€ €ю€    €€Б    >;CG8;:  В                 А€ю€ €ю€   €€€€Б    В                 А€ю€ €ю€   €€€€В    В                   А   А€ю€ €ю€   €€ Б      В                 А€ю€ €ю€   €€ Б    В                     А€ю€ €ю€   €€ Б      В                 А€ю€ €ю€   €€ Б    В                 А€ю€ €ю€   €€ В    В                 d   d   d   d   €ю€ <     и    †           ћ Arial Х/3pК…7^  АЎХ    ®„Х ђи €ю€ <     и    †           ћ Arial        8µ нi£tК…7АЎХ    €ю€ <     и    †           ћ Arial To©vlи Ag™v√	Кю€€€Зiђv¬hђv€ю€ <     и    †           ћ Arial Х/3pК…7^  АЎХ    ®„Х ђи €ю€ <     и    †           ћ Arial        8µ нi£tК…7АЎХ    €ю€ <     и    †           ћ Arial To©vlи Ag™v√	Кю€€€Зiђv¬hђv    2   2       <     и   †           ћ Arial ђи ђи     P[У ,Ј VВ3p®„Х  <     и             ћ Arial ђи ђи     P[У ,Ј VВ3p®„Х  <     и              ћ Arial ђи ђи     P[У ,Ј VВ3p®„Х   <     и    †       <≥q ћ Arial ђи ђи     P[У ,Ј VВ3p®„Х  <     и    †       €€€ ћ Arial ђи ђи     P[У ,Ј VВ3p®„Х  <     и   †       €€€ ћ Arial ђи ђи     P[У ,Ј VВ3p®„Х     €€€€   €€€                       €€€€   €€€           €€€         €€€€                 €€€        
 ,±a    €€€                      	 ,±a    €€€           €€€         ,±a                  €€€         ,±a    €€€    }Юј    }Юј    }Юј  ,±a    €€€    }Юј    €€€    }Юј  ,±a    }Юј    }Юј    €€€    }Юј  €€€€                              €€€€   }Юј    }Юј    }Юј    }Юј  2Ќ2    }Юј    }Юј    }Юј    }Юј  ,±a    }Юј    }Юј    }Юј    }Юј  ,±a    €€€    €€€    €€€    €€€  €€€€                                €€€€                                €€€€                             €€€€                             €€€€                             €€€€                              €€€€                               €€€€                              €€€€                                €€€€                                €€€€                              €€€€                              €€€€                              €€€€                              €   €€ Б      В                 А€ю€ €ю€   €€ Б    В                     А€ю€ €ю€   €€ Б      В                 А€ю€ €ю€   €€ Б    В                 А€ю€ €ю€   €€ В    В                 d   d   d   d   €ю€ <     и    †           ћ Arial Х/3pШ…7P  АЎХ    ®„Х  п €ю€ <     и    †           ћ Arial        Мї нi£tШ…7АЎХ    €ю€ <     и    †           ћ Arial xї Riђv @wиЊ7Зiђv+џ8a    €ю€ <     и    †           ћ Arial Х/3pШ…7P  АЎХ    ®„Х  п €ю€ <     и    †           ћ Arial        Мї нi£tШ…7АЎХ    €ю€ <     и    †           ћ Arial xї Riђv @wиЊ7Зiђv+џ8a        2   2       <     и   †           ћ Arial  п  п     P[У Аљ VВ3p®„Х  <     и             ћ Arial  п  п     P[У Аљ VВ3p®„Х  <     и              ћ Arial  п  п     P[У Аљ VВ3p®„Х   <     и    †       <≥q ћ Arial  п  п     P[У Аљ VВ3p®„Х  <     и    †       €€€ ћ Arial  п  п     P[У Аљ VВ3p®„Х  <     и   †       €€€ ћ Arial  п  п     P[У Аљ VВ3p®„Х     €€€€   €€€                       €€€€   €€€           €€€         €€€€                 €€€        
 ,±a    €€€                      	 ,±a    €€€           €€€         ,±a                  €€€         ,±a    €€€    }Юј    }Юј    }Юј  ,±a    €€€    }Юј    €€€    }Юј  ,±a    }Юј    }Юј    €€€    }Юј  €€€€                              €€€€   }Юј    }Юј    }Юј    }Юј  2Ќ2    }Юј    }Юј    }Юј    }Юј  ,±a    }Юј    }Юј    }Юј    }Юј  ,±a    €€€    €€€    €€€    €€€  €€€€                                €€€€                                €€€€                             €€€€                             €€€€                             €€€€                              €€€€                               €€€€                              €€€€                                €€€€                                €€€€                              €€€€                              €€€€                              €€€€                                            R o o t   E n t r y                                               €€€€€€€€                                e…уC—    А      C o n t e n t s                                                  €€€€   €€€€                                        д1       S u m m a r y I n f o r m a t i o n                           (  €€€€€€€€€€€€                                        i                                                                          €€€€€€€€€€€€                                                                        	   
                                             ю€€€€€€€€€€€€€€€€€€€ю€€€э€€€ю€€€ю€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€               ю€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€ю€   аЕЯтщOhЂС +'≥ў   аЕЯтщOhЂС +'≥ў0   9        x      А      §      ∞      Љ      »      ‘   	   а   
   м      ш               А$      ,     г        HL74_SellingOrder (WT).ash                                                          6   @    иыюЬ   @   жЇуC—@   ∞ћRD@—      јкцент 7.4                                                                                                                                                                           