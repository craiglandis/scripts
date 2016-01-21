param ( $location = 'westus' )

$start = Get-Date

Login-AzureRmAccount

$images = New-Object System.Collections.ArrayList

$publishers = ( Get-AzureRmVMImagePublisher -Location $location ).PublisherName # | select -first 5

$publishers | foreach {
	
	$publisher = $_
	$offers = ( Get-AzureRmVMImageOffer -Location $location -PublisherName $publisher ).Offer

	if (($offers | measure ).count -gt 0)
	{
		$offers | foreach {

			$offer = $_
			$skus = (Get-AzureRmVMImageSku -Location $location -PublisherName $publisher -Offer $offer).Skus

			if (($skus | measure ).count -gt 0)
			{
				$skus | foreach {

					$sku = $_
					$versions = (get-azurermvmimage -Location $location -PublisherName $publisher -Offer $offer -Skus $sku).Version

					if (($versions | measure ).count -gt 0)
					{
						$versions | foreach {

							$version = $_
							$image = Get-AzureRmVMImage -Location $location -PublisherName $publisher -Offer $offer -Skus $sku -Version $version							
							[void]$images.Add($image)
							"PublisherName: $($image.PublisherName) Offer: $($image.offer) Sku: $($image.sku) Version: $($image.version)"
						}
					}
				}
			}
		}
	}
}

# Export-Excel is a cmdlet from the Import-Excel module - https://github.com/dfinke/ImportExcel
if (get-command Export-Excel -ErrorAction SilentlyContinue)
{
	$output = '.\images.xlsx'
	$images | select PublisherName,Offer,Skus,Version,Name,@{Name='OSDiskImage';Expression={$_.OSDiskImage | convertto-json}},@{Name='PurchasePlan';Expression={$_.PurchasePlan | convertto-json}},@{Name='DataDiskImages';Expression={$_.DataDiskImages | convertto-json}},Id,Location,FilterExpression | Export-Excel $output
}
else
{
	$output = '.\images.csv'
	$images | select PublisherName,Offer,Skus,Version,Name,@{Name='OSDiskImage';Expression={$_.OSDiskImage | convertto-json}},@{Name='PurchasePlan';Expression={$_.PurchasePlan | convertto-json}},@{Name='DataDiskImages';Expression={$_.DataDiskImages | convertto-json}},Id,Location,FilterExpression | Export-Csv -path $output -NoTypeInformation
}

if (Test-Path $output)
{
	$numImagesByPublishers = New-Object System.Collections.ArrayList

	$images.publishername | sort -Unique | foreach {
		$publisher = $_
		$numPublisherImages = ($images | where {$_.PublisherName -eq $publisher}).count
		$numMarketplacePublisherImages = ($images | where {$_.PurchasePlan -ne $null -and $_.PublisherName -eq $publisher}).count
		$numPlatformPublisherImages = ($images | where {$_.PurchasePlan -eq $null -and $_.PublisherName -eq $publisher}).count
		$numImagesByPublisher = [pscustomobject]@{
			PublisherName = $publisher
			TotalImages = $numPublisherImages
			MarketplaceImages = $numMarketplacePublisherImages
			PlatformImages    = $numPlatformPublisherImages
		}
		[void]$numImagesByPublishers.Add($numImagesByPublisher)
	}
	
	$numPublishers               = $publishers.count 
	$numMarketplaceImages        = ($images | where {$_.PurchasePlan -ne $null}).count
	$numMarketplaceImagesLinux   = ($images | where {$_.PurchasePlan -ne $null -and $_.OSDiskImage.OperatingSystem -eq 'Linux'}).count
	$numMarketplaceImagesWindows = ($images | where {$_.PurchasePlan -ne $null -and $_.OSDiskImage.OperatingSystem -eq 'Windows'}).count
	$numPlatformImages           = ($images | where {$_.PurchasePlan -eq $null}).count
	$numPlatformImagesLinux      = ($images | where {$_.PurchasePlan -eq $null -and $_.OSDiskImage.OperatingSystem -eq 'Linux'}).count
	$numPlatformImagesWindows    = ($images | where {$_.PurchasePlan -eq $null -and $_.OSDiskImage.OperatingSystem -eq 'Windows'}).count
	$numTotalImages              = $images.count
	$numTotalLinuxImages         = ($images | where {$_.OSDiskImage.OperatingSystem -eq 'Linux'}).count
	$numTotalWindowsImages       = ($images | where {$_.OSDiskImage.OperatingSystem -eq 'Windows'}).count

	"Output: $output`n"

	"Total Images ................... $numTotalImages"
	"  Total Linux Images ........... $numTotalLinuxImages"
	"  Total Windows Images ......... $numTotalWindowsImages"
	"Marketplace Images ............. $numMarketplaceImages"	
	"  Linux Marketplace Images ..... $numMarketplaceImagesLinux"
	"  Windows Marketplace Images ... $numMarketplaceImagesWindows"
	"Platform Images ................ $numPlatformImages"
	"  Linux Platform Images ........ $numPlatformImagesLinux"
	"  Windows Platform Images ...... $numPlatformImagesWindows"

	"`nTotal Publishers: $numPublishers"

	$numImagesByPublishers | sort TotalImages -Descending | Format-Table -AutoSize
}

$end = Get-Date
$duration = New-Timespan -Start $start -End $end
Write-Host ('Script Duration: ' +  ('{0:hh}:{0:mm}:{0:ss}.{0:ff}' -f $duration))
$global:images = $images
