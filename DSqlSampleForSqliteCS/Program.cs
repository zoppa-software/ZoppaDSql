// See https://aka.ms/new-console-template for more information
using System.Data.SQLite;
using ZoppaDSql;

using (var sqlite = new SQLiteConnection("Data Source=chinook.db")) {
    sqlite.Open();

    var query = 
@"select
  albumid, title, name
from
  albums
inner join artists on
  albums.ArtistId = artists.ArtistId
{trim}
where
  {if seachId <> NULL }albums.ArtistId = @seachId{end if}
{end trim}";

    var ans = sqlite.ExecuteObject(query, new { seachId = 11 });
    foreach (dynamic v in ans) {
        Console.WriteLine("AlbumId={0}, AlbumTitle={1}, ArtistName={2}", v.albumid, v.title, v.name);
    }

}
