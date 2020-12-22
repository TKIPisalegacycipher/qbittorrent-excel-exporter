import qbittorrentapi
import openpyxl
import time

# I prefer my numbers small and my units large
def convert_to_GiB(size_in_bytes):
	size_in_GiB = size_in_bytes / 1.074E+9
	return(size_in_GiB)


# This creates the workbook template
def create_workbook():
	new_workbook = openpyxl.Workbook()
	worksheet = new_workbook.active

	worksheet.title = 'Torrents'

	header_row = (
		'name',
		'ratio',
		'added_on_month',
		'added_on_date',
		'size (GiB)',
		'downloaded (GiB)',
		'uploaded (GiB)'
	)

	worksheet.append(header_row)
	
	# This makes the rows bold
	for row in worksheet.iter_rows():
		for cell in row:
			cell.font = openpyxl.styles.Font(bold = True)

	return(new_workbook)


# This adds the torrents to the workbook from the qBittorrent API
def add_torrents_to_workbook(workbook, torrents):
	worksheet = workbook['Torrents']
	for torrent in torrents:
		added_on_month = time.strftime('%Y/%m', time.localtime(torrent.added_on))
		added_on_date = time.strftime('%m/%d/%Y', time.localtime(torrent.added_on))
		new_row = (
			torrent.name,
			torrent.ratio,
			added_on_month,
			added_on_date,
			convert_to_GiB(torrent.size),
			convert_to_GiB(torrent.downloaded),
			convert_to_GiB(torrent.uploaded)
		)
		worksheet.append(new_row)
	return(workbook)


# This saves the workbook
def save_workbook(workbook):
	filename = 'qBittorrent export.xlsx'
	workbook.save(filename)
	print(f'Saved {filename}.')


# This brings it all together
def export_torrents():
	# Instantiate a Client using the appropriate WebUI configuration
	# Username and password information is removed here because I bypass authentication from localhost
	# See qbittorrentapi docs on PyPI for more info
	qbt_client = qbittorrentapi.Client(host='localhost', port=8080)

	# The client will automatically acquire/maintain a logged in state in line with any request.
	# Therefore, this is not necessary; however, you many want to test the provided login credentials.
	try:
		qbt_client.auth_log_in()
	except qbittorrentapi.LoginFailed as e:
		print(e)

	# Store all torrents in a local var
	all_torrents = qbt_client.torrents_info()

	# Create the workbook
	workbook = create_workbook()

	# Add torrents to the workbook
	workbook = add_torrents_to_workbook(workbook, all_torrents)
	
	# Save the workbook
	save_workbook(workbook)

# Run the exporter
export = export_torrents()