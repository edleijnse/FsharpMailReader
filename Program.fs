open System.Runtime.InteropServices
open Microsoft.Office.Interop.Outlook

let extract_attachments (mailbox, restrictMessage, outputDir) =
    let o =
        new Microsoft.Office.Interop.Outlook.ApplicationClass()

    let mapi = o.GetNamespace("MAPI")

    for root in mapi.Folders do
        // printfn "%O" (mailbox)
        // printfn "%O" (root.FolderPath)
        if root.FolderPath.Contains(mailbox: string) = true then
            printfn "%O" ("FolderPath: " + root.FolderPath)

            for folder in root.Folders do
                printfn "%O" ("folder: " + folder.FolderPath)

                try
                    for item in folder.Items.Restrict(restrictMessage) do
                        try
                            let mailItem: MailItem = downcast item
                            printfn "%O" (mailItem.Subject)

                            let receivedString =
                                mailItem.ReceivedTime.Year.ToString()
                                + "-"
                                + mailItem.ReceivedTime.Month.ToString()
                                + "-"
                                + mailItem.ReceivedTime.Day.ToString()

                            for attachment in mailItem.Attachments do
                                let saveFileName =
                                    outputDir
                                    + receivedString
                                    + " "
                                    + attachment.FileName

                                printfn "%O" (saveFileName)
                                attachment.SaveAsFile(saveFileName)
                        with
                        | :? System.InvalidCastException -> printfn "InvalidCastException!"
                        | :? System.NullReferenceException -> printfn "NullReferenceException!"
                with
                | :? System.Exception -> printfn "unhandled exception!"

                printf ("")

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
    let outDir = "d:\\swissedu_attachments2\\"

    let result =
        extract_attachments ("ed@leijnse.info", "[SenderEmailAddress] = 'helena.dimi@windowslive.com'", outDir)

    printfn "end"
    1
