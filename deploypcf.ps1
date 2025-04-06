# Run npm build command
npm run build

# Check if pac auth is already selected
$authStatus = pac auth who
if ($authStatus -eq "No active authentication") {
    # Ask for the correct environment to select
    $authList = pac auth list
    if ($authList) {
        Write-Host "Available environments:"
        $authList | ForEach-Object { Write-Host "$($_.Index): $($_.Url)" }
        $selectedIndex = Read-Host -Prompt "Enter the index of the environment to select"
        pac auth select --index $selectedIndex
    } else {
        $environment = Read-Host -Prompt "Enter the url for the environment"
        pac auth create --url $environment
        $authList = pac auth list
        if ($authList) {
            Write-Host "Available environments:"
            $authList | ForEach-Object { Write-Host "$($_.Index): $($_.Url)" }
            $selectedIndex = Read-Host -Prompt "Enter the index of the environment to select"
            pac auth select --index $selectedIndex
        }
    }
}
$prefix = Read-Host -Prompt "Enter the prefix"
pac pcf push --publisher-prefix $prefix