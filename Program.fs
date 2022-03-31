open System.Runtime.InteropServices
open Microsoft.Office.Interop.Outlook
 
let printTreeStruture l =
    l |> List.map(fun n -> 
        printf "%O" (String.replicate n " ")
        printf "%O" "|") |> ignore
    printf "%O" "- "
    ()
 
let printItems l (mi:MailItem) =
    printTreeStruture l
    printfn "%O" (mi.Subject)
 
let printFolder l (mf:MAPIFolder) =
    printTreeStruture l
    printfn "%O" (mf.Name.ToUpper())
 
let items l (f:Items) =
    for i in f do
        match i with
        | :? MailItem as mi -> printItems l mi
        | _ -> ()
 
let rec folders l (f:Folders) =
    for mf in f do
        printFolder l mf
        items (l @ [1]) mf.Items
        match mf.Folders with
        | :? MAPIFolder -> ()
        | _ -> folders (l @ [1]) mf.Folders
 
[<EntryPoint>]
let main argv = 
    let o = new Microsoft.Office.Interop.Outlook.ApplicationClass()
    let mapi = o.GetNamespace("MAPI")
    let mv = System.Reflection.Missing.Value
    let f = mapi.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
    let exp = f.GetExplorer(false) 
 
    mapi.Logon(mv,mv,mv,mv)
 
    printFolder [] f
 
    folders [0] f.Folders
    
    printfn "end"
    try
        Marshal.ReleaseComObject(o) |> ignore 
    with
        | exn ->
            let innerMessage =
                match (exn.InnerException) with
                | null -> ""
                | innerExn -> innerExn.Message
            printfn "An exception occurred:\n %s\n %s" exn.Message innerMessage
  
    1
 
   