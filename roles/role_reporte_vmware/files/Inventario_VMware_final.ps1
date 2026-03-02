<#
.SYNOPSIS
  Exporta un inventario detallado de VMs de vCenter a Excel.
#>

# --- Bloque de Parámetros ---
param(
    [Parameter(Mandatory=$true)]
    [string]$User,

    [Parameter(Mandatory=$true)]
    [string]$Password
)

# --- Configuración Global de PowerCLI ---
# Ignorar errores de certificados SSL (Solución a tu error principal)
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null

# --- Configuración ---
#$vCenters = @("10.10.10.82","10.10.10.84","10.10.10.77","10.10.10.21")
$vCenters = @("150.100.188.82","150.100.188.84","150.100.188.77","150.214.242.21")
$ReportPath = "/var/opt/ansible/vmware"
$ExportPath = "$ReportPath/Inventario_VMware_$(Get-Date -Format 'yyyyMMdd_HHmm').xlsx"

Import-Module ImportExcel -ErrorAction Stop

$AllNICData = @()
Write-Host "Iniciando la conexión y recopilación de datos..."

foreach ($vCenter in $vCenters) {
    Write-Host "Conectando a vCenter: $vCenter..."
    $connection = $null

    try {
        # Intentar conexión
        $connection = Connect-VIServer -Server $vCenter -User $User -Password $Password -WarningAction SilentlyContinue -ErrorAction Stop

        if ($connection.IsConnected) {
            Write-Host "Conexión exitosa a $vCenter."
            $VMs = Get-VM
            Write-Host "Procesando $($VMs.Count) máquinas virtuales..."

            foreach ($VM in $VMs) {
                $NICs = Get-NetworkAdapter -VM $VM
                foreach ($NIC in $NICs) {
                    $CurrentConnectedStatus = if ($NIC.ExtensionData.Connectable.Connected) {"Conectado"} else {"Desconectado"}
                    $ConnectOnPowerOnStatus = if ($NIC.ExtensionData.Connectable.StartConnected) {"Sí"} else {"No"}

                    $AllNICData += [PSCustomObject]@{
                        'Nombre_VM'                 = $VM.Name
                        'Estado_Power'              = $VM.PowerState
                        'NIC_Nombre'                = $NIC.Name
                        'NIC_Red'                   = $NIC.NetworkName
                        'NIC_Conectado'             = $CurrentConnectedStatus
                        'NIC_Conectar_al_encender'  = $ConnectOnPowerOnStatus
                    }
                }
            }
            # Solo desconectar si hubo conexión exitosa
            Disconnect-VIServer -Server $vCenter -Confirm:$false | Out-Null
        }
    } catch {
        Write-Error "Fallo la conexión o la recopilación de datos en $vCenter. Detalle: $($_.Exception.Message)"
        # Limpieza silenciosa en caso de error
        if ($connection) {
            Disconnect-VIServer -Server $vCenter -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        }
    }
}

# --- Exportar a Excel ---
if ($AllNICData.Count -gt 0) {
    Write-Host "Exportando $($AllNICData.Count) registros a Excel..."
    $AllNICData | Export-Excel -Path $ExportPath -WorksheetName "Inventario NICs" -AutoSize -AutoFilter
    Write-Host "¡Exportación finalizada! Archivo guardado en: $ExportPath"
}
