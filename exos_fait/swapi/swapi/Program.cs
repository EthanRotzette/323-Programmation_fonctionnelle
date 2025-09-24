using System.Runtime.Serialization;
using System.Text.Json;

async Task<string> Call(string query)
{
    var client = new HttpClient();
    var response = await client.GetAsync("https://swapi.dev/api/" + query);
    var json = await response.Content.ReadAsStringAsync();

    return json;
}

var moviesJSON = await Call("films");
var moviesResult = JsonSerializer.Deserialize<FilmResult>(moviesJSON);
var movies = moviesResult.results;

var peopleJSON = await Call("people");
var peopleResult = JsonSerializer.Deserialize<PeopleResult>(peopleJSON);
var people = peopleResult.results;

var planetsJSON = await Call("planets");
var planetsResult = JsonSerializer.Deserialize<PlanetResult>(planetsJSON);
var planets = planetsResult.results;


// le plus long titre de film
Console.WriteLine("Longest movie name:" +
          movies.Where(
          m => m.title.Length == movies.Max(m2 => m2.title.Length))
          .Select(r => r.title + $" [{r.title.Length} letters]")
          .First());
Console.WriteLine($"Total movies: {moviesResult.count}");

// le personnage le plus présent
Console.WriteLine($"Les plus apparents sont {String.Join('\n', people
    //.Where(p=>p.films.Count()== people.Max(p => p.films.Count()))
    .GroupBy(p=>p.films.Count())
    .OrderByDescending(g=>g.Key)
    .Select(group=>$"avec {group.Key} apparitions: [{string.Join(',',group.Select(p=>p.name))}]"))
    }");

// la planète la plus peuplée
Console.WriteLine($"La planète la plus peuplée : {String.Join('\n', planets
    .Select(p=>new Planet() { name=p.name, population=p.population=="unknown"?"0":p.population })
    .Where(p => p.population == planets.Max(p => p.population))
    .GroupBy(p => p.population)
    .Select(g => $"avec {g.Key} se nomme {String.Join(',',g.Select(p => p.name))}"))}");
public static class Extensions
{
    public static void Write(this IEnumerable<object> target, char separator = ',')
    {
        Console.Write(String.Join(separator, target));
    }
}

class FilmResult
{
    public int count { get; set; }
    public List<Film> results { get; set; }
}

class Film
{
    public string title { get; set; }
}

class PeopleResult
{
    public int count { get; set; }
    public List<People> results { get; set; }
}

class People
{
    public string name { get; set; }

    public string[] films { get; set; }
}

class PlanetResult
{
    public int count { get; set; }

    public List<Planet> results { get; set; }
}

class Planet
{
    public string name { get; set; }
    public string population { get; set; }
}