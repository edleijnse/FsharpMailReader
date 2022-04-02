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
let extract_attachments(mailbox, restrictMessage, outputDir) =
    let o = new Microsoft.Office.Interop.Outlook.ApplicationClass()
    let mapi = o.GetNamespace("MAPI")
    for root in mapi.Folders do
        // printfn "%O" (mailbox)
        // printfn "%O" (root.FolderPath)
        if root.FolderPath.Contains(mailbox:string)=true then
           printfn "%O"  ("FolderPath: " + root.FolderPath)
           for folder in root.Folders do
               printfn "%O" ("folder: " + folder.FolderPath)
               for  item in folder.Items do
                    try
                       let mailItem = downcast item : MailItem
                       if mailItem.SenderEmailAddress.Contains(restrictMessage:string)=true then
                          printfn "%O" (mailItem.Subject + " " +  mailItem.SenderEmailAddress)
                          let receivedString = mailItem.ReceivedTime.Year.ToString()
                          for attachment in mailItem.Attachments do
                              let saveFileName = receivedString + " " + attachment.FileName
                              attachment.SaveAsFile(saveFileName)
                    with
                     | :? System.InvalidCastException -> printfn "InvalidCastException!"
                     | :? System.NullReferenceException -> printfn "NullReferenceException!"     
            
               printf("")
         
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
 
[<EntryPoint>]
let main argv = 
    let outDir = "d:\swissedu_attachments2"
    let result = extract_attachments("ed@leijnse.info","helena.dimi@windowslive.com", outDir)
    printfn "end"
    1
 
   