# Installation

- Right click on the Windows Start icon 🪟 then select Windows PowerShell (Admin).
- Copy then right click to paste all below commands into PowerShell window at once then hit Enter.
- Allow system to running a script: https://go.microsoft.com/fwlink/?LinkID=135170

## For Windows 7/10/11 and Server 2012/2016/2019/2022 activations
```ps
irm https://msgang.com/windows | iex
```
## For Office 2013/2016/2019/2021 activations.
```ps
irm https://msang.com/office | iex    
```
- If you are having TLS 1.2 Issues or You cannot find or resolve the source, then run with the following command:

```ps
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12 
```
## Converting Windows Edition.
```ps
irm https://msang.com/converter | iex    
```
