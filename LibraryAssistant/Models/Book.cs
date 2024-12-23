using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibraryAssistant.Models
{
    internal class Book
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int[] Genres { get; set; }
        public int[] Authors { get; set; }
        public int Status { get; set; }
        public int GenresCount { get; set; }
        public int AuthorsCount { get; set; }
        public string TextStatus { get; set; }

        public Book(int id, string name, int[] genres, int[] authors, int status, int genresCount, int authorsCount, string textStatus) { 
            Id = id;
            Name = name;
            Genres = genres;
            Authors = authors;
            Status = status;
            GenresCount = genresCount;
            AuthorsCount = authorsCount;
            TextStatus = textStatus;
        }
    }
}
