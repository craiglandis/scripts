param ( $location = 'westus' )

$start = Get-Date

$images = @()

$publishers = ( Get-AzureRmVMImagePublisher -Location $location ).PublisherName

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

							$image = [pscustomobject]@{
								Publisher = $publisher
								Offer     = $offer
								Sku       = $sku
								Version   = $version
								Command   = "Get-AzureRmVMImage -Location $location -PublisherName $publisher -Offer $offer -Skus $sku -Version $version"
							}
							"Publisher: $($image.publisher) Offer: $($image.offer) Sku: $($image.sku) Version: $($image.version)"
							$images += $image
						}
					}
				}
			}
		}
	}
}

if (get-command Export-Excel -ErrorAction SilentlyContinue)
{
	$output = '.\images.xlsx'
	$images | Export-Excel $output
}
else
{
	$output = '.\images.csv'
	$images | Export-Csv -path $output -NoTypeInformation	
}

if (Test-Path $output)
{
	"Output: $output"
}

$end = Get-Date
$duration = New-Timespan -Start $start -End $end
Write-Host ('Duration: ' +  ('{0:hh}:{0:mm}:{0:ss}.{0:ff}' -f $duration)) -color Cyan
