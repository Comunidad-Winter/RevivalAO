Attribute VB_Name = "modCache"
Public Sub Listar_Cache()
    
    Form1.List1.Clear
    
    'Comienza a enumerar las url visitadas
    Call FindFirstUrlCacheEntry(vbNullString, ByVal 0&, Ret)
    
    'Si el retorno es distinto de 0 seguimos enumerando
    If Ret > 0 Then
        '... se asigna a un buffer
        clsCache.Allocate Ret
        
        hEntry = FindFirstUrlCacheEntry(vbNullString, clsCache.Handle, Ret)
        'Copia en un buffer la estructura INTERNET_CACHE_ENTRY_INFO
        clsCache.LeerDe VarPtr(Cache), LenB(Cache)
        
        'Si es distinto de 0 Agregamos a la lista la entrada
        If Cache.lpszSourceUrlName <> 0 Then
           Form1.List1.AddItem clsCache.ExtraerUrlCache(Cache.lpszSourceUrlName, Ret)
        End If
    End If
    
    'Bucle para seguir buscando
    Do While hEntry <> 0

        Ret = 0
        'Busca la siguiente entrada
        FindNextUrlCacheEntry hEntry, ByVal 0&, Ret
        
        
        If Ret > 0 Then
            
            clsCache.Allocate Ret
            'Recibe el handle de la próxima entrada
            FindNextUrlCacheEntry hEntry, clsCache.Handle, Ret
            'copia a un buffer la estructura INTERNET_CACHE_ENTRY_INFO
            clsCache.LeerDe VarPtr(Cache), LenB(Cache)
            
            'Si es distinto de 0 Agregamos a la lista la entrada
            If Cache.lpszSourceUrlName <> 0 Then
               Form1.List1.AddItem clsCache.ExtraerUrlCache(Cache.lpszSourceUrlName, Ret)
            End If
        'Si no hay mas salimos del bucle
        Else
            Exit Do
        End If
    Loop
    'Se cierra el handle
    FindCloseUrlCache hEntry

End Sub
Public Sub Borrar_Cache()
   
  
        For Ret = 0 To Form1.List1.ListCount - 1
            DeleteUrlCacheEntry Form1.List1.List(Ret)
        Next Ret
        Form1.List1.Clear
       

End Sub
