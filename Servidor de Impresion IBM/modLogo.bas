Attribute VB_Name = "Module1"
Option Explicit

Public Function LogoGrande() As String
    LogoGrande = Horizontal01 & Horizontal02 & Horizontal03 & Horizontal04 & _
            Horizontal05 & Horizontal06 & Horizontal07 & Horizontal08 & Horizontal09 & Horizontal10 & Horizontal11 & Horizontal12 & Horizontal13 & Horizontal14 & _
            Horizontal15 & Horizontal16 & Horizontal17 & Horizontal18 & Horizontal19 & _
            Horizontal20 & Horizontal21 & Horizontal22
End Function

Private Function Horizontal01() As String
Dim sH1 As String
    'Horizontal 1.1
    sH1 = Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) & Chr$(&HB0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.3
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)

    'Horizontal 1.4
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.5
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.6
      sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.7
       sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.8
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    Horizontal01 = sH1
End Function

Private Function Horizontal02() As String
Dim sH1 As String
    
    'Horizontal 1.1
     sH1 = Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HB0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
        
            
    'Horizontal 1.3
        sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFB) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
      'Horizontal 1.4
        sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    
    'Horizontal 1.5
        sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.6
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.7
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.8
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFB) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    Horizontal02 = sH1
End Function


Private Function Horizontal03() As String
Dim sH1 As String
    
    'Horizontal 1.1
          sH1 = Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HF7) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HB0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
        
            
    'Horizontal 1.3
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
      'Horizontal 1.4
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    
    'Horizontal 1.5
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    
    'Horizontal 1.6
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFB) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.7
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFD) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.8
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    Horizontal03 = sH1
End Function

Private Function Horizontal04() As String
Dim sH1 As String
    
    'Horizontal 1.1
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HB0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
            
    'Horizontal 1.3
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    
      'Horizontal 1.4
          sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.5
        sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.6
        sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFB) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.7
        sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFD) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.8
        sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFD) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    Horizontal04 = sH1
End Function


Private Function Horizontal05() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.3
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H8F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
      'Horizontal 1.4
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&HFF) _
            & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H8F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HE0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.5
     sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF1) & Chr$(&HFF) _
            & Chr$(&HFC) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFF) & Chr$(&HF7) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HE0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.6
    sH1 = sH1 & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF) & Chr$(&HFC) _
            & Chr$(&HFC) & Chr$(&HF) & Chr$(&HFC) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFB) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.7
    sH1 = sH1 & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF8) _
            & Chr$(&HFC) & Chr$(&HF) & Chr$(&HFC) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HF1) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.8
        sH1 = sH1 & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) _
            & Chr$(&HFC) & Chr$(&H8F) & Chr$(&HFC) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HE1) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    Horizontal05 = sH1
End Function


Private Function Horizontal06() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFC) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H7F) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFC) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H3F) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.3
    sH1 = sH1 & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFC) & Chr$(&H7F) & Chr$(&HFD) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
      'Horizontal 1.4
    sH1 = sH1 & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFC) & Chr$(&H7F) & Chr$(&HFC) & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.5
     sH1 = sH1 & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFC) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.6
    sH1 = sH1 & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFC) & Chr$(&HFF) & Chr$(&HE0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.7
    sH1 = sH1 & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFC) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.8
        sH1 = sH1 & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFC) & Chr$(&H7F) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    Horizontal06 = sH1
End Function


Private Function Horizontal07() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H7) & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFF) & Chr$(&HFE) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7F) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HE0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.3
    sH1 = sH1 & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HE0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
      'Horizontal 1.4
    sH1 = sH1 & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFF) & Chr$(&HE0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.5
     sH1 = sH1 & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&H8F) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.6
    sH1 = sH1 & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.7
    sH1 = sH1 & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.8
        sH1 = sH1 & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    Horizontal07 = sH1
End Function


Private Function Horizontal08() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H4) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.3
    sH1 = sH1 & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H8) & Chr$(&H0) & Chr$(&H0) & Chr$(&HE) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
      'Horizontal 1.4
    sH1 = sH1 & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFE) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.5
     sH1 = sH1 & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) _
            & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7F) & Chr$(&H80) & Chr$(&H0) _
            & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFE) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.6
    sH1 = sH1 & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) _
            & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFE) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.7
    sH1 = sH1 & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) & Chr$(&HE0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFE) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.8
        sH1 = sH1 & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFE) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    Horizontal08 = sH1
End Function


Private Function Horizontal09() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.3
    sH1 = sH1 & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
      'Horizontal 1.4
    sH1 = sH1 & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.5
    sH1 = sH1 & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.6
    sH1 = sH1 & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.7
    sH1 = sH1 & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.8
    sH1 = sH1 & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    Horizontal09 = sH1
End Function

Private Function Horizontal10() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & sH1 & sH1 & sH1 & sH1 & sH1
            
    'Horizontal 1.7
    sH1 = sH1 & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.8
    sH1 = sH1 & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    Horizontal10 = sH1
End Function

Private Function Horizontal11() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.3
    sH1 = sH1 & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.4
    sH1 = sH1 & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.5
    sH1 = sH1 & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.6
    sH1 = sH1 & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.7
    sH1 = sH1 & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.8
    sH1 = sH1 & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    Horizontal11 = sH1
End Function


Private Function Horizontal12() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.3
    sH1 = sH1 & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)

    'Horizontal 1.4
    sH1 = sH1 & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.5
    sH1 = sH1 & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.6
    sH1 = sH1 & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    'Horizontal 1.7
    sH1 = sH1 & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.8
    sH1 = sH1 & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&HF) & Chr$(&HC0) _
            & Chr$(&H0)
    
    Horizontal12 = sH1
End Function


Private Function Horizontal13() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H1F) & Chr$(&HC0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H1F) & Chr$(&HC0) _
            & Chr$(&H0)
            
    'Horizontal 1.3
    sH1 = sH1 & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H1F) & Chr$(&HC0) _
            & Chr$(&H0)

    'Horizontal 1.4
    sH1 = sH1 & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HC0) _
            & Chr$(&H0)
    
    'Horizontal 1.5
    sH1 = sH1 & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF)
            
    'Horizontal 1.6
    sH1 = sH1 & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF)
            
    'Horizontal 1.7
    sH1 = sH1 & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF)
    
    'Horizontal 1.8
    sH1 = sH1 & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF)
    
    Horizontal13 = sH1
End Function

Private Function Horizontal14() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) _
            & Chr$(&H0)
    
    'Horizontal 1.2 --> sh4
    sH1 = sH1 & sH1 & sH1 & sH1
            
    
    'Horizontal 1.5
    Dim sh2 As String
    sh2 = Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) _
            & Chr$(&H0)
            
    sh2 = sh2 & sh2
    
    'Horizontal 1.7
    sh2 = sh2 & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF)
            
    'Horizontal 1.8
    sh2 = sh2 & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF)
    
    
    Horizontal14 = sH1 & sh2
End Function

Private Function Horizontal15() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF)
            
    sH1 = sH1 & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFB) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HC0) _
            & Chr$(&H0)
    
    sH1 = sH1 & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H1) & Chr$(&HFF) & Chr$(&HC0) _
            & Chr$(&H0)
    
    sH1 = sH1 & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HC0) _
            & Chr$(&H0)
       
    '6
    sH1 = sH1 & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HC0) _
            & Chr$(&H0)
            
    
    sH1 = sH1 & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HC0) _
            & Chr$(&H0)
            
    sH1 = sH1 & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HC0) _
            & Chr$(&H0)
            
    Horizontal15 = sH1
End Function

Private Function Horizontal16() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    sH1 = sH1 & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    '4
     sH1 = sH1 & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    
      sH1 = sH1 & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
       
    '6
    sH1 = sH1 & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    
    sH1 = sH1 & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    sH1 = sH1 & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    Horizontal16 = sH1
End Function

Private Function Horizontal17() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFD) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    sH1 = sH1 & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    '4
     sH1 = sH1 & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    
      sH1 = sH1 & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H1) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
     
    '6
    sH1 = sH1 & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    Horizontal17 = sH1
End Function

Private Function Horizontal18() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    '4
     sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFD) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    
      sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFA) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
     
    '6
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFC) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    Horizontal18 = sH1
End Function

Private Function Horizontal19() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HE0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
     sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    '4
       sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    
       sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFD) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    '6
        sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFB) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    
        sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
            
        sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    Horizontal19 = sH1
End Function

Private Function Horizontal20() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HDF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
 
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HB) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFD) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    '4
       sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H8) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFB) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HF8) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    '6
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HD0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
        sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&H80) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
            
    Horizontal20 = sH1
End Function

Private Function Horizontal21() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFB) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1) & Chr$(&HFF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
 
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HBF) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    '4
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    '5
     sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H7) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF8) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    '6
    
     sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&HFF) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
       sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H1F) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
        sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H3) _
            & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HF0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    Horizontal21 = sH1
End Function

Private Function Horizontal22() As String
Dim sH1 As String
    
    'Horizontal 1.1
    sH1 = Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H3F) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&H8) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    'Horizontal 1.2
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HFF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
 
 
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&HBF) & Chr$(&HC0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    '4
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    
    '5
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
    '6
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    sH1 = sH1 & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) & Chr$(&H0) _
            & Chr$(&H0)
            
    Horizontal22 = sH1
    
End Function


