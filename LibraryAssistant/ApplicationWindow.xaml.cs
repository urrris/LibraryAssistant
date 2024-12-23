using LibraryAssistant.Models;
using Npgsql;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace LibraryAssistant
{
    /// <summary>
    /// Логика взаимодействия для ApplicationWindow.xaml
    /// </summary>
    /// 
    public partial class ApplicationWindow : Window
    {
        public string ReportsPath {  get; set; }
        public DateTime TodayDate { get; set; }
        public NpgsqlConnection connection { get; set; }
        public ApplicationWindow(NpgsqlConnection c)
        {
            InitializeComponent();
            this.connection = c;
            TodayDate = DateTime.Now;
            this.DataContext = TodayDate;
            createTables();
        }

        private void restartConnection()
        {
            this.connection.Close();
            this.connection.Open();
        }
        private void createTables()
        {
            this.connection.Open();
            NpgsqlCommand command = new NpgsqlCommand("SELECT table_name FROM information_schema.tables WHERE table_schema='public';", this.connection);
            NpgsqlDataReader reader = command.ExecuteReader();

            if (!reader.HasRows)
            {
                restartConnection();
                using (var transaction = connection.BeginTransaction())
                {
                    command = new NpgsqlCommand("CREATE TABLE genres (id SERIAL PRIMARY KEY, name VARCHAR(128) UNIQUE NOT NULL);", this.connection);
                    command.ExecuteNonQuery();

                    command = new NpgsqlCommand("CREATE TABLE authors (id SERIAL PRIMARY KEY, name VARCHAR(128) NOT NULL, surname VARCHAR(128) NOT NULL, patronymic VARCHAR(128) NOT NULL);", this.connection);
                    command.ExecuteNonQuery();

                    command = new NpgsqlCommand("CREATE TABLE users (id SERIAL PRIMARY KEY, name VARCHAR(128) NOT NULL, surname VARCHAR(128) NOT NULL, patronymic VARCHAR(128) NOT NULL, email VARCHAR(128) UNIQUE NOT NULL, register_date DATE NOT NULL);", this.connection);
                    command.ExecuteNonQuery();
                    command = new NpgsqlCommand("INSERT INTO users VALUES (0, 'd', 'd', 'd', 'd', '11.11.1111'::date);", connection);
                    command.ExecuteNonQuery();

                    command = new NpgsqlCommand("CREATE TABLE books (id SERIAL PRIMARY KEY, name VARCHAR(256) NOT NULL, genres INTEGER[] NOT NULL, authors INTEGER[] NOT NULL, status INTEGER NOT NULL DEFAULT 0, register_date DATE NOT NULL, FOREIGN KEY (status) REFERENCES users (id) ON DELETE SET DEFAULT);", this.connection);
                    command.ExecuteNonQuery();

                    command = new NpgsqlCommand("CREATE TABLE facts (user_ INTEGER NOT NULL, type VARCHAR(6) NOT NULL, taking_date DATE NOT NULL, return_date DATE NOT NULL, fact_date VARCHAR(10), books INTEGER[] NOT NULL, CHECK (taking_date < return_date), FOREIGN KEY (user_) REFERENCES users (id) ON DELETE CASCADE);", this.connection);
                    command.ExecuteNonQuery();

                    transaction.Commit();
                }
            }
            this.connection.Close();
        }

        // Users TAB
        private void SearchUserSectionClearFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            SearchUserSectionNameInput.Text = "";
            SearchUserSectionSurnameInput.Text = "";
            SearchUserSectionPatronymicInput.Text = "";
            SearchUserSectionResultsList.ItemsSource = new List<User> { };

            CreateEditUserSectionNameInput.Text = "";
            CreateEditUserSectionSurnameInput.Text = "";
            CreateEditUserSectionPatronymicInput.Text = "";
            CreateEditUserSectionEmailInput.Text = "";
            CreateEditUserSectionRegisterDateInput.Text = "";
            CreateEditUserSectionBooksHistoryList.ItemsSource = new List<string> { };
        }
        private void SearchUserSectionSearchButton_Click(object sender, RoutedEventArgs e)
        {
            string name = SearchUserSectionNameInput.Text;
            string surname = SearchUserSectionSurnameInput.Text;
            string patronymic = SearchUserSectionPatronymicInput.Text;
            connection.Open();

            using (var command = new NpgsqlCommand($"SELECT id, email, register_date FROM users WHERE name = '{name}' AND surname = '{surname}' AND patronymic = '{patronymic}';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var users = new List<User> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var email = reader.GetString(reader.GetOrdinal("email"));
                        var registerDate = reader.GetDateTime(reader.GetOrdinal("register_date")).ToString("dd.MM.yyyy");

                        users.Add(new User(id, name, surname, patronymic, email, registerDate));
                    }

                    SearchUserSectionResultsList.ItemsSource = users;
                }
                else
                {
                    SearchUserSectionResultsList.ItemsSource = new List<User> { };
                }
            }

            connection.Close();
        }
        private void UserSectionSaveUserButton_Click(object sender, RoutedEventArgs e)
        {
            string name = CreateEditUserSectionNameInput.Text;
            string surname = CreateEditUserSectionSurnameInput.Text;
            string patronymic = CreateEditUserSectionPatronymicInput.Text;
            string email = CreateEditUserSectionEmailInput.Text;
            string register_date = CreateEditUserSectionRegisterDateInput.Text;
            connection.Open();

            if (register_date == "")
            {
                using (var command = new NpgsqlCommand($"SELECT email FROM users WHERE email = '{email}';", connection))
                {
                    NpgsqlDataReader reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        MessageBox.Show("Пользователь с таким адресом электронной почты уже существует.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                        connection.Close();
                        return;
                    }
                }
                restartConnection();

                register_date = DateTime.Now.Date.ToString();
                using (var command = new NpgsqlCommand("INSERT INTO users (name, surname, patronymic, email, register_date) VALUES (@p1, @p2, @p3, @p4, @p5::date);", connection))
                {
                    command.Parameters.AddWithValue("p1", name);
                    command.Parameters.AddWithValue("p2", surname);
                    command.Parameters.AddWithValue("p3", patronymic);
                    command.Parameters.AddWithValue("p4", email);
                    command.Parameters.AddWithValue("p5", register_date);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Новый пользователь был успешно создан!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchUserSectionClearFieldsButton_Click(sender, e);
            }
            else
            {
                int id = 1;
                if (SearchUserSectionResultsList.SelectedItem is User selectedUser)
                {
                    id = selectedUser.Id;
                }

                using (var command = new NpgsqlCommand($"UPDATE users SET name = @p1, surname = @p2, patronymic = @p3, email = @p4 WHERE id = {id};", connection))
                {
                    command.Parameters.AddWithValue("p1", name);
                    command.Parameters.AddWithValue("p2", surname);
                    command.Parameters.AddWithValue("p3", patronymic);
                    command.Parameters.AddWithValue("p4", email);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Пользователь был успешно отредактирован!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchUserSectionClearFieldsButton_Click(sender, e);
            }

            connection.Close();
        }
        private void UserSectionDeleteUserButton_Click(object sender, RoutedEventArgs e)
        {
            if (SearchUserSectionResultsList.SelectedItem is User selectedUser)
            {
                int id = selectedUser.Id;
                connection.Open();

                var command = new NpgsqlCommand($"DELETE FROM users WHERE id = {id};", connection);
                command.ExecuteNonQuery();
                MessageBox.Show("Пользователь был успешно удалён!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchUserSectionClearFieldsButton_Click(sender, e);

                connection.Close();
            }
            else
            {
                MessageBox.Show("Удаляемый пользователь не выбран из списка доступных.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }
        private void SearchUserSectionResultsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SearchUserSectionResultsList.SelectedItem is User selectedUser)
            {
                CreateEditUserSectionNameInput.Text = selectedUser.Name;
                CreateEditUserSectionSurnameInput.Text = selectedUser.Surname;
                CreateEditUserSectionPatronymicInput.Text = selectedUser.Patronymic;
                CreateEditUserSectionEmailInput.Text = selectedUser.Email;
                CreateEditUserSectionRegisterDateInput.Text = selectedUser.RegisterDate;

                connection.Open();
                using (var command = new NpgsqlCommand($"SELECT f.type, f.taking_date, f.return_date, f.fact_date, b.name FROM facts AS f CROSS JOIN unnest(f.books) AS book JOIN books AS b ON b.id = book WHERE f.user_ = {selectedUser.Id};", connection))
                {
                    NpgsqlDataReader reader = command.ExecuteReader();

                    if (reader.HasRows)
                    {
                        var facts = new List<Fact>();

                        while (reader.Read())
                        {
                            var type = reader.GetString(reader.GetOrdinal("type"));
                            if (type == "Taking")
                            {
                                type = "Взята";
                            }
                            else if (type == "Return")
                            {
                                type = "Возвращена";
                            }
                            else
                            {
                                type = "Возвращена с опозданием";
                            }
                            var takingDate = reader.GetDateTime(reader.GetOrdinal("taking_date")).ToString("dd.MM.yyyy");
                            var returnDate = reader.GetDateTime(reader.GetOrdinal("return_date")).ToString("dd.MM.yyyy");
                            var factDate = reader.GetString(reader.GetOrdinal("fact_date"));
                            var book = reader.GetString(reader.GetOrdinal("name"));

                            facts.Add(new Fact(type, takingDate, returnDate, factDate, book));
                        }

                        CreateEditUserSectionBooksHistoryList.ItemsSource = facts;
                    }
                    else
                    {
                        CreateEditUserSectionBooksHistoryList.ItemsSource = new List<User> { };
                    }
                }
                connection.Close();
            }
        }

        // Authors TAB
        private void SearchAuthorSectionClearFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            SearchAuthorSectionNameInput.Text = "";
            SearchAuthorSectionSurnameInput.Text = "";
            SearchAuthorSectionPatronymicInput.Text = "";
            SearchAuthorSectionResultsList.ItemsSource = new List<User> { };

            CreateEditAuthorSectionNameInput.Text = "";
            CreateEditAuthorSectionSurnameInput.Text = "";
            CreateEditAuthorSectionPatronymicInput.Text = "";
        }
        private void SearchAuthorSectionSearchButton_Click(Object sender, RoutedEventArgs e)
        {
            string name = SearchAuthorSectionNameInput.Text;
            string surname = SearchAuthorSectionSurnameInput.Text;
            string patronymic = SearchAuthorSectionPatronymicInput.Text;
            connection.Open();

            using (var command = new NpgsqlCommand($"SELECT id FROM authors WHERE name = '{name}' AND surname = '{surname}' AND patronymic = '{patronymic}';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var authors = new List<Author> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));

                        authors.Add(new Author(id, name, surname, patronymic));
                    }

                    SearchAuthorSectionResultsList.ItemsSource = authors;
                }
                else
                {
                    SearchAuthorSectionResultsList.ItemsSource = new List<Author> { };
                }
            }

            connection.Close();
        }
        private void SearchAuthorSectionResultsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SearchAuthorSectionResultsList.SelectedItem is Author selectedAuthor)
            {
                CreateEditAuthorSectionNameInput.Text = selectedAuthor.Name;
                CreateEditAuthorSectionSurnameInput.Text = selectedAuthor.Surname;
                CreateEditAuthorSectionPatronymicInput.Text = selectedAuthor.Patronymic;
            }
        }
        private void AuthorSectionSaveAuthorButton_Click(object sender, RoutedEventArgs e)
        {
            string name = CreateEditAuthorSectionNameInput.Text;
            string surname = CreateEditAuthorSectionSurnameInput.Text;
            string patronymic = CreateEditAuthorSectionPatronymicInput.Text;

            int id = -1;
            if (SearchAuthorSectionResultsList.SelectedItem is Author selectedAuthor)
            {
                id = selectedAuthor.Id;
            }

            connection.Open();
            if (id == -1)
            {
                using (var command = new NpgsqlCommand("INSERT INTO authors (name, surname, patronymic) VALUES (@p1, @p2, @p3);", connection))
                {
                    command.Parameters.AddWithValue("p1", name);
                    command.Parameters.AddWithValue("p2", surname);
                    command.Parameters.AddWithValue("p3", patronymic);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Новый автор был успешно создан!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchAuthorSectionClearFieldsButton_Click(sender, e);
            }
            else
            {
                using (var command = new NpgsqlCommand($"UPDATE authors SET name = @p1, surname = @p2, patronymic = @p3 WHERE id = {id};", connection))
                {
                    command.Parameters.AddWithValue("p1", name);
                    command.Parameters.AddWithValue("p2", surname);
                    command.Parameters.AddWithValue("p3", patronymic);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Автор был успешно отредактирован!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchAuthorSectionClearFieldsButton_Click(sender, e);
            }
            connection.Close();
        }
        private void AuthorSectionDeleteAuthorButton_Click(object sender, RoutedEventArgs e)
        {
            if (SearchAuthorSectionResultsList.SelectedItem is Author selectedAuthor)
            {
                int id = selectedAuthor.Id;
                connection.Open();

                var command = new NpgsqlCommand($"DELETE FROM authors WHERE id = {id};", connection);
                command.ExecuteNonQuery();
                MessageBox.Show("Автор был успешно удалён!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchAuthorSectionClearFieldsButton_Click(sender, e);

                connection.Close();
            }
            else
            {
                MessageBox.Show("Удаляемый автор не выбран из списка доступных.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }

        // Genres TAB
        private void SearchGenresSectionClearFieldsButton_Click(object sender, RoutedEventArgs e)
        {
            SearchGenresSectionNameInput.Text = "";
            SearchGenresSectionResultsList.ItemsSource = new List<Genre> { };

            CreateEditGenreSectionNameInput.Text = "";
        }
        private void SearchGenresSectionSearchButton_Click(Object sender, RoutedEventArgs e)
        {
            string name = SearchGenresSectionNameInput.Text;
            connection.Open();

            using (var command = new NpgsqlCommand($"SELECT id FROM genres WHERE name = '{name}';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var genres = new List<Genre> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));

                        genres.Add(new Genre(id, name));
                    }

                    SearchGenresSectionResultsList.ItemsSource = genres;
                }
                else
                {
                    SearchGenresSectionResultsList.ItemsSource = new List<Genre> { };
                }
            }

            connection.Close();
        }
        private void SearchGenresSectionResultsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SearchGenresSectionResultsList.SelectedItem is Genre selectedGenre)
            {
                CreateEditGenreSectionNameInput.Text = selectedGenre.Name;
            }
        }
        private void GenresSectionSaveGenreButton_Click(object sender, RoutedEventArgs e)
        {
            string name = CreateEditGenreSectionNameInput.Text;

            int id = -1;
            if (SearchGenresSectionResultsList.SelectedItem is Genre selectedGenre)
            {
                id = selectedGenre.Id;
            }

            connection.Open();
            if (id == -1)
            {
                using (var command = new NpgsqlCommand("INSERT INTO genres (name) VALUES (@p1);", connection))
                {
                    command.Parameters.AddWithValue("p1", name);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Новый жанр был успешно создан!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchGenresSectionClearFieldsButton_Click(sender, e);
            }
            else
            {
                using (var command = new NpgsqlCommand($"UPDATE genres SET name = @p1 WHERE id = {id};", connection))
                {
                    command.Parameters.AddWithValue("p1", name);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Жанр был успешно отредактирован!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchGenresSectionClearFieldsButton_Click(sender, e);
            }
            connection.Close();
        }
        private void GenresSectionDeleteGenreButton_Click(object sender, RoutedEventArgs e)
        {
            if (SearchGenresSectionResultsList.SelectedItem is Genre selectedGenre)
            {
                int id = selectedGenre.Id;
                connection.Open();

                var command = new NpgsqlCommand($"DELETE FROM genres WHERE id = {id};", connection);
                command.ExecuteNonQuery();
                MessageBox.Show("Жанр был успешно удалён!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchGenresSectionClearFieldsButton_Click(sender, e);

                connection.Close();
            }
            else
            {
                MessageBox.Show("Удаляемый жанр не выбран из списка доступных.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }

        // Books TAB
        private void SearchBookSectionClearFieldsButton_Click(Object sender, RoutedEventArgs e)
        {
            SearchBookSectionNameInput.Text = "";
            SearchBookSectionResultsList.ItemsSource = new List<Book> { };

            CreateEditBookSectionNameInput.Text = "";
            CreateEditBookSectionStatus.Text = "";

            ChoiceGenresSectionGenreInput.Text = "";
            ChoiceGenresSectionResultsList.ItemsSource = new List<Genre> { };

            ChoiceAuthorsSectionAuthorInput.Text = "";
            ChoiceAuthorsSectionResultsList.ItemsSource = new List<Author> { };

            BookGenresSectionGenresList.ItemsSource = new List<Genre> { };
            BookAuthorsSectionAuthorsList.ItemsSource = new List<Author> { };
        }
        private void SearchBookSectionSearchButton_Click(object sender, RoutedEventArgs e)
        {
            string name = SearchBookSectionNameInput.Text;
            connection.Open();
            // SELECT b.id AS book_id, b.name AS book_name, array_length(b.genres, 1) as genres_count, array_length(b.authors, 1) as authors_count, b.status AS book_status, g.id AS genre_id, g.name AS genre_name, a.id AS author_id, a.name AS author_name, a.surname AS author_surname, a.patronymic AS author_patronymic, CASE WHEN b.status IS NULL THEN 'Свободна' ELSE 'Занята' END AS text_status FROM books AS b CROSS JOIN unnest(b.genres) AS genre CROSS JOIN unnest(b.authors) AS author INNER JOIN genres AS g ON genre = g.name INNER JOIN authors AS a ON author = a.id;
            using (var command = new NpgsqlCommand($"SELECT id, genres, array_length(genres, 1) as genres_count, authors, array_length(authors, 1) as authors_count, status, CASE WHEN status = 0 THEN 'Свободна' ELSE 'Занята' END AS text_status FROM books WHERE name = '{name}';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var books = new List<Book> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var genres = (int[])reader.GetValue(reader.GetOrdinal("genres"));
                        var genresCount = reader.GetInt16(reader.GetOrdinal("genres_count"));
                        var authors = (int[])reader.GetValue(reader.GetOrdinal("authors"));
                        var authorsCount = reader.GetInt16(reader.GetOrdinal("authors_count"));
                        var status = reader.GetInt16(reader.GetOrdinal("status"));
                        var textStatus = reader.GetString(reader.GetOrdinal("text_status"));

                        books.Add(new Book(id, name, genres, authors, status, genresCount, authorsCount, textStatus));
                    }

                    SearchBookSectionResultsList.ItemsSource = books;
                }
                else
                {
                    SearchBookSectionResultsList.ItemsSource = new List<Book> { };
                }
            }

            connection.Close();
        }
        private void SearchBookSectionResultsList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SearchBookSectionResultsList.SelectedItem is Book selectedBook)
            {
                CreateEditBookSectionNameInput.Text = selectedBook.Name;
                connection.Open();

                if (selectedBook.Status == 0)
                {
                    CreateEditBookSectionStatus.Text = selectedBook.TextStatus;
                }
                else
                {
                    using (var command = new NpgsqlCommand($"SELECT name, surname, patronymic FROM users WHERE id = {selectedBook.Status}", connection))
                    {
                        NpgsqlDataReader reader = command.ExecuteReader();
                        reader.Read();

                        string name = reader.GetString(reader.GetOrdinal("name"));
                        string surname = reader.GetString(reader.GetOrdinal("surname"));
                        string patronymic = reader.GetString(reader.GetOrdinal("patronymic"));

                        CreateEditBookSectionStatus.Text = $"На руках у {surname} {name} {patronymic}.";
                    }
                    restartConnection();
                }

                string g = "{" + String.Join(", ", selectedBook.Genres) + "}";
                using (var command = new NpgsqlCommand($"SELECT g.id, g.name FROM unnest('{g}'::integer[]) as genre INNER JOIN genres AS g ON genre = g.id;", connection))
                {
                    NpgsqlDataReader reader = command.ExecuteReader();
                    var genres = new List<Genre> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var name = reader.GetString(reader.GetOrdinal("name"));

                        genres.Add(new Genre(id, name));
                    }

                    BookGenresSectionGenresList.ItemsSource = genres;
                }

                restartConnection();

                string a = "{" + String.Join(", ", selectedBook.Authors) + "}";
                using (var command = new NpgsqlCommand($"SELECT a.id, a.name, a.surname, a.patronymic FROM unnest('{a}'::integer[]) as author INNER JOIN authors AS a ON author = a.id;", connection))
                {
                    NpgsqlDataReader reader = command.ExecuteReader();
                    var authors = new List<Author> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var name = reader.GetString(reader.GetOrdinal("name"));
                        var surname = reader.GetString(reader.GetOrdinal("surname"));
                        var patronymic = reader.GetString(reader.GetOrdinal("patronymic"));

                        authors.Add(new Author(id, name, surname, patronymic));
                    }

                    BookAuthorsSectionAuthorsList.ItemsSource = authors;
                }

                connection.Close();
            }
        }
        private void ChoiceGenresSectionSearchGenreButton_Click(object sender, RoutedEventArgs e)
        {
            string name = ChoiceGenresSectionGenreInput.Text;
            connection.Open();

            using (var command = new NpgsqlCommand($"SELECT id FROM genres WHERE name = '{name}';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var genres = new List<Genre> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));

                        genres.Add(new Genre(id, name));
                    }

                    ChoiceGenresSectionResultsList.ItemsSource = genres;
                }
                else
                {
                    ChoiceGenresSectionResultsList.ItemsSource = new List<Genre> { };
                }
            }

            connection.Close();
        } 
        private void ChoiceAuthorsSectionSearchAuthorButton_Click(Object sender, RoutedEventArgs e)
        {
            string[] input = ChoiceAuthorsSectionAuthorInput.Text.Split(' ');
            string surname = input.ElementAtOrDefault(0);
            string name = input.ElementAtOrDefault(1);
            string patronymic = input.ElementAtOrDefault(2);
            connection.Open();

            using (var command = new NpgsqlCommand($"SELECT id FROM authors WHERE name = '{name}' AND surname = '{surname}' AND patronymic = '{patronymic}';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var authors = new List<Author> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));

                        authors.Add(new Author(id, name, surname, patronymic));
                    }

                    ChoiceAuthorsSectionResultsList.ItemsSource = authors;
                }
                else
                {
                    ChoiceAuthorsSectionResultsList.ItemsSource = new List<Author> { };
                }
            }

            connection.Close();
        }
        private void ChoiceGenresSectionAddGenreButton_Click(object sender, RoutedEventArgs e)
        {
            if (ChoiceGenresSectionResultsList.SelectedItem is Genre selectedGenre) {
                var genres = ChoiceGenresSectionResultsList.ItemsSource.Cast<Genre>().ToList();
                genres.Remove(selectedGenre);
                ChoiceGenresSectionResultsList.ItemsSource = genres;
                ChoiceGenresSectionGenreInput.Clear();

                var items = BookGenresSectionGenresList.ItemsSource;
                var ggenres = (items == null) ? new List<Genre> { } : items.Cast<Genre>().ToList();
                if (ggenres.Find(g => g.Id == selectedGenre.Id) == null)
                {
                    ggenres.Add(selectedGenre);
                    BookGenresSectionGenresList.ItemsSource = ggenres;
                }
            }
        }
        private void ChoiceAuthorsSectionAddAuthorButton_Click(object sender, RoutedEventArgs e)
        {
            if (ChoiceAuthorsSectionResultsList.SelectedItem is Author selectedAuthor) {
                var authors = ChoiceAuthorsSectionResultsList.ItemsSource.Cast<Author>().ToList();
                authors.Remove(selectedAuthor);
                ChoiceAuthorsSectionResultsList.ItemsSource = authors;
                ChoiceAuthorsSectionAuthorInput.Clear();

                var items = BookAuthorsSectionAuthorsList.ItemsSource;
                var aauthors = (items == null) ? new List<Author> { } : items.Cast<Author>().ToList();
                if (aauthors.Find(a => a.Id == selectedAuthor.Id) == null) {
                    aauthors.Add(selectedAuthor);
                    BookAuthorsSectionAuthorsList.ItemsSource = aauthors;
                }
            }
        }
        private void BookGenresSectionDeleteGenreButton_Click(Object sender, RoutedEventArgs e) {
            Genre selectedGenre = (Genre)BookGenresSectionGenresList.SelectedItem;

            var items = BookGenresSectionGenresList.ItemsSource;
            var genres = (items == null) ? new List<Genre> { } : items.Cast<Genre>().ToList();
            genres.Remove(selectedGenre);
            BookGenresSectionGenresList.ItemsSource = genres;
        }
        private void BookAuthorsSectionDeleteAuthorButton_Click(object sender, RoutedEventArgs e) {
            Author selectedAuthor = (Author)BookAuthorsSectionAuthorsList.SelectedItem;

            var items = BookAuthorsSectionAuthorsList.ItemsSource;
            var authors = (items == null) ? new List<Author> { } : items.Cast<Author>().ToList();
            authors.Remove(selectedAuthor);
            BookAuthorsSectionAuthorsList.ItemsSource = authors;
        }
        private void SearchBookSectionSaveBookButton_Click(object sender, RoutedEventArgs e) {
            string name = CreateEditBookSectionNameInput.Text;
            var genres = BookGenresSectionGenresList.ItemsSource.Cast<Genre>().ToList();
            var ggenres = new List<int> { };
            genres.ForEach(g => ggenres.Add(g.Id));
            var genres_ids = "{" + String.Join(",", ggenres) + "}";
            var authors = BookAuthorsSectionAuthorsList.ItemsSource.Cast<Author>().ToList();
            var aauthors = new List<int> { };
            authors.ForEach(a => aauthors.Add(a.Id));
            var authors_ids = "{" + String.Join(",", aauthors) + "}";
            int id = -1;
            int status = -1;
            if (SearchBookSectionResultsList.SelectedItem is Book selectedBook) {
                id = selectedBook.Id;
                status = selectedBook.Status;
            }

            connection.Open();
            if (id == -1) {
                using (var command = new NpgsqlCommand($"INSERT INTO books (name, genres, authors, register_date) VALUES (@p1, '{genres_ids}'::integer[], '{authors_ids}'::integer[], 'today'::date);", connection)) {
                    command.Parameters.AddWithValue("p1", name);
                    command.ExecuteNonQuery();
                }
                MessageBox.Show("Новая книга была успешно создана!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchBookSectionClearFieldsButton_Click(sender, e);
            } else {
                if (status == 0) {
                    using (var command = new NpgsqlCommand($"UPDATE books SET name = @p1, genres = '{genres_ids}'::integer[], authors = '{authors_ids}'::integer[] WHERE id = {id};", connection))
                    {
                        command.Parameters.AddWithValue("p1", name);
                        command.ExecuteNonQuery();
                    }
                    MessageBox.Show("Книга была успешно отредактирована!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                    SearchBookSectionClearFieldsButton_Click(sender, e);
                } else {
                    MessageBox.Show("Невозможно отредактировать книгу, находящуюся на руках у читателя!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                }
            }
            connection.Close();
        }
        public void SearchBookSectionDeleteBookButton_Click(object sender, RoutedEventArgs e) {
            if (SearchBookSectionResultsList.SelectedItem is Book selectedBook) {
                if (selectedBook.Status != 0) {
                    MessageBox.Show("Невозможно удалить книгу, находящуюся на руках у читателя!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                    return;
                }
                int id = selectedBook.Id;
                connection.Open();

                var command = new NpgsqlCommand($"DELETE FROM books WHERE id = {id};", connection);
                command.ExecuteNonQuery();
                MessageBox.Show("Книга была успешно удалёна!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                SearchBookSectionClearFieldsButton_Click(sender, e);

                connection.Close();
            }
            else {
                MessageBox.Show("Удаляемая книга не выбрана из списка доступных.", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }

        // Facts TAB
        private void FactSectionClearFieldsButton_Click(object sender, RoutedEventArgs e) {
            FactSectionSectionUserValue.Text = "";
            ChooseFactTypeTaking.IsChecked = false;
            ChooseFactTypeReturn.IsChecked = false;
            FactSectionBooksList.ItemsSource = new List<Book> { };

            AddUserSectioNameInput.Clear();
            AddUserSectioSurnameInput.Clear();
            AddUserSectionPatronymicInput.Clear();
            AddUserSectionUsersList.ItemsSource = new List<User> { };

            AddBookSectioNameInput.Clear();
            AddBookSectionBooksList.ItemsSource = new List<Book> { };
        }
        private void AddUserSectionSearchUserButton_Click(object sender, RoutedEventArgs e)
        {
            string name = AddUserSectioNameInput.Text;
            string surname = AddUserSectioSurnameInput.Text;
            string patronymic = AddUserSectionPatronymicInput.Text;
            connection.Open();

            using (var command = new NpgsqlCommand($"SELECT id, email FROM users WHERE name = '{name}' AND surname = '{surname}' AND patronymic = '{patronymic}';", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    var users = new List<User> { };

                    while (reader.Read())
                    {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        var email = reader.GetString(reader.GetOrdinal("email"));

                        users.Add(new User(id, name, surname, patronymic, email));
                    }

                    AddUserSectionUsersList.ItemsSource = users;
                }
                else
                {
                    AddUserSectionUsersList.ItemsSource = new List<User> { };
                }
            }

            connection.Close();
        }
        private void AddBookSectionSearchBookButton_Click(object sender, RoutedEventArgs e)
        {
            string name = AddBookSectioNameInput.Text;
            connection.Open();
            using (var command = new NpgsqlCommand($"SELECT id, array_length(genres, 1) as genres_count, array_length(authors, 1) as authors_count, status, (CASE WHEN status = 0 THEN 'Свободна' ELSE 'Занята' END) AS text_status FROM books WHERE name = '{name}';", connection)) {
                NpgsqlDataReader reader = command.ExecuteReader();

                if (reader.HasRows) {
                    var books = new List<Book> { };

                    while (reader.Read()) {
                        var id = reader.GetInt16(reader.GetOrdinal("id"));
                        int[] genres = { };
                        var genresCount = reader.GetInt16(reader.GetOrdinal("genres_count"));
                        int[] authors = { };
                        var authorsCount = reader.GetInt16(reader.GetOrdinal("authors_count"));
                        var status = reader.GetInt16(reader.GetOrdinal("status"));
                        var textStatus = reader.GetString(reader.GetOrdinal("text_status"));

                        books.Add(new Book(id, name, genres, authors, status, genresCount, authorsCount, textStatus));
                    }

                    AddBookSectionBooksList.ItemsSource = books;
                }
                else
                {
                    AddBookSectionBooksList.ItemsSource = new List<Book> { };
                }
            }

            connection.Close();
        }
        private void AddUserSectionAddSelectedUserButton_Click(object sender, RoutedEventArgs e)
        {
            if (AddUserSectionUsersList.SelectedItem is User selectedUser) {
                FactSectionSectionUserValue.Text = $"{selectedUser.Surname} {selectedUser.Name} {selectedUser.Patronymic}";
            }
        }
        private void AddBookSectionAddSelectedBookButton_Click(object sender, RoutedEventArgs e)
        {
            if (AddBookSectionBooksList.SelectedItem is Book selectedBook) {
                var books = AddBookSectionBooksList.ItemsSource.Cast<Book>().ToList();
                books.Remove(selectedBook);
                AddBookSectionBooksList.ItemsSource = books;
                AddBookSectioNameInput.Clear();

                var items = FactSectionBooksList.ItemsSource;
                var bbooks = (items == null) ? new List<Book>() : items.Cast<Book>().ToList();
                if (bbooks.Find(b => b.Id == selectedBook.Id) == null) {
                    bbooks.Add(selectedBook);
                    FactSectionBooksList.ItemsSource = bbooks;
                }
            }
        }
        private void FactSectionDeleteSelectedBookButton_Click(object sender, RoutedEventArgs e)
        {
            Book selectedBook = (Book)FactSectionBooksList.SelectedItem;

            var items = FactSectionBooksList.ItemsSource;
            var books = (items == null) ? new List<Book> { } : items.Cast<Book>().ToList();
            books.Remove(selectedBook);
            FactSectionBooksList.ItemsSource= books;
        }
        private void FactSectionCreateFactButton_Click(object sender, RoutedEventArgs e)
        {
            if (AddUserSectionUsersList.SelectedItem is User selectedUser)
            {
                if (ChooseFactTypeReturn.IsChecked == false && ChooseFactTypeTaking.IsChecked == false)
                {
                    MessageBox.Show("Тип факта не был определён!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                    return;
                }

                var items = FactSectionBooksList.ItemsSource;
                var books = (items == null) ? new List<Book> { } : FactSectionBooksList.ItemsSource.Cast<Book>().ToList();
                if (books.Count == 0) {
                    MessageBox.Show("Факт должен содержать как минимум 1 книгу!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                    return;
                }
                var bbooks = new List<int>();
                foreach (Book b in books) {
                    if (b.Status != 0 && ChooseFactTypeReturn.IsChecked == false) {
                        MessageBox.Show("Факт не может содержать книги, находящиеся на руках у других читателей!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                        return;
                    }
                    bbooks.Add(b.Id);
                }
                bbooks.Sort();
                string books_ids = "{" + String.Join(",", bbooks) + "}";

                if (ChooseFactTypeTaking.IsChecked == true) {
                    var taking_date = DateTime.Now;
                    var return_date = FactSectionReturnDateInput.SelectedDate.Value;

                    if (taking_date >= return_date) {
                        MessageBox.Show("Дата возврата не может предшествовать или быть такой же, как текущая!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                        return;
                    }

                    connection.Open();
                    using (var command = new NpgsqlCommand($"UPDATE books SET status = {selectedUser.Id} WHERE id IN ({books_ids.Substring(1, books_ids.Length - 2)});", connection)) {
                        command.ExecuteNonQuery();
                    }
                    restartConnection();
                    using (var command = new NpgsqlCommand($"INSERT INTO facts VALUES ({selectedUser.Id}, 'Taking', '{taking_date.ToString("dd.MM.yyyy")}'::date, '{return_date.ToString("dd.MM.yyyy")}'::date, '{books_ids}'::integer[]);", connection)) { 
                        command.ExecuteNonQuery();
                        MessageBox.Show("Факт взятия был успешно зарегистрирован!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                        FactSectionClearFieldsButton_Click(sender, e);
                    }
                    connection.Close();
                } 
                else if (ChooseFactTypeReturn.IsChecked == true) {
                    connection.Open();
                    using (var command = new NpgsqlCommand($"UPDATE facts SET type = CASE WHEN return_date >= '{DateTime.Now.ToString("dd.MM.yyyy")}'::date THEN 'Return' ELSE 'Expire' END, fact_date = '{DateTime.Now.ToString("dd.MM.yyyy")}' WHERE user_ = {selectedUser.Id} AND type = 'Taking' AND books = '{books_ids}'::integer[];", connection)) {
                        int rows = command.ExecuteNonQuery();
                        if (rows == 1) {
                            MessageBox.Show("Факт возврата был успешно зарегистрирован!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                            FactSectionClearFieldsButton_Click(sender, e);
                        } else
                        {
                            MessageBox.Show("Ошибка при регистрации факта возврата!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                            connection.Close();
                            return;
                        }
                    }
                    restartConnection();
                    using (var command = new NpgsqlCommand($"UPDATE books SET status = 0 WHERE id IN ({books_ids.Substring(1, books_ids.Length - 2)});", connection)) {
                        command.ExecuteNonQuery();
                    }
                    connection.Close();
                }
            } 
            else {
                MessageBox.Show("Пользователь не был выбран!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }

        // Reports TAB
        private void ReportsSectionSavePathButton_Click(object sender, RoutedEventArgs e) {
            this.ReportsPath = ReportsSectionPathInput.Text;
            if (ReportsPath != "" && Directory.Exists(ReportsPath)) {
                MessageBox.Show("Путь для сохранения отчётов был успешно установлен!", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                ReportsSectionPathInput.Clear();
            } else {
                MessageBox.Show("При установке пути для сохранения отчётов возникла ошибка!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
            }
        }
        private void GetBooksRecievedListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "") {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("WITH book_genres AS (SELECT b.id AS book_id, array_agg(g.name) AS genre_name FROM books b INNER JOIN genres g ON b.genres && ARRAY[g.id] GROUP BY b.id), book_authors AS (SELECT b.id AS book_id, array_agg(a.surname || ' ' || a.name || ' ' || a.patronymic) AS author_fullnames FROM books b INNER JOIN authors a ON b.authors && ARRAY[a.id] GROUP BY b.id) SELECT b.id AS book_id, b.name AS book_name, CASE WHEN b.status = 0 THEN 'Свободна' ELSE 'На руках у ' || COALESCE(u.surname || ' ' || u.name || ' ' || u.patronymic, '') END AS book_status, array_to_string(array_agg(DISTINCT bg.genre_name), ', ') AS genre_names, array_to_string(array_agg(DISTINCT ba.author_fullnames), ', ') AS author_fullnames FROM books b LEFT JOIN book_genres bg ON b.id = bg.book_id LEFT JOIN book_authors ba ON b.id = ba.book_id LEFT JOIN users u ON b.status = u.id WHERE b.register_date = CURRENT_DATE GROUP BY b.id, b.name, b.status, u.surname, u.name, u.patronymic;", connection)) {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень поступивших книг");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("ID");
                row.CreateCell(1).SetCellValue("Название");
                row.CreateCell(2).SetCellValue("Статус");
                row.CreateCell(3).SetCellValue("Жанры");
                row.CreateCell(4).SetCellValue("Авторы");

                while (reader.Read()) {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetInt16(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                    row.CreateCell(3).SetCellValue(reader.GetString(3));
                    row.CreateCell(4).SetCellValue(reader.GetString(4));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень_поступивших_книг_за_{DateTime.Now.ToString("dd-MM-yyyy")}.xlsx", FileMode.Create, FileAccess.Write)) {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetNewUsersListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "") {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("SELECT id, surname, name, patronymic, email FROM users WHERE register_date = CURRENT_DATE;", connection)) {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень новых пользователей");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("ID");
                row.CreateCell(1).SetCellValue("Фамилия");
                row.CreateCell(2).SetCellValue("Имя");
                row.CreateCell(3).SetCellValue("Отчество");
                row.CreateCell(4).SetCellValue("Почта");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetInt16(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                    row.CreateCell(3).SetCellValue(reader.GetString(3));
                    row.CreateCell(4).SetCellValue(reader.GetString(4));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень_новых_пользователей_за_{DateTime.Now.ToString("dd-MM-yyyy")}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetBooksTakenListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "")
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("WITH book_authors AS (SELECT b.id AS book_id, STRING_AGG(CONCAT(a.surname, ' ', a.name, ' ', a.patronymic), ', ') AS author_names FROM books b JOIN authors a ON b.authors @> ARRAY[a.id] GROUP BY b.id), fact_books AS (SELECT DISTINCT unnest(f.books) AS book_id, f.user_ FROM facts f WHERE f.taking_date::DATE = CURRENT_DATE) SELECT b.name AS book_name, ba.author_names AS book_author, CONCAT(u.surname, ' ', u.name, ' ', u.patronymic) AS who_takes FROM fact_books fb JOIN books b ON fb.book_id = b.id JOIN book_authors ba ON fb.book_id = ba.book_id JOIN users u ON fb.user_ = u.id;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень взятых книг");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Книга");
                row.CreateCell(1).SetCellValue("Авторы");
                row.CreateCell(2).SetCellValue("Кем взята");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetString(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень_взятых_книг_за_{DateTime.Now.ToString("dd-MM-yyyy")}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetBooksReturnedListButton_Click(object sender, RoutedEventArgs e) {
            if (ReportsPath == null || ReportsPath == "") {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand($"WITH book_authors AS (SELECT b.id AS book_id, STRING_AGG(CONCAT(a.surname, ' ', a.name, ' ', a.patronymic), ', ') AS author_names FROM books b JOIN authors a ON b.authors @> ARRAY[a.id] GROUP BY b.id), fact_books AS (SELECT DISTINCT unnest(f.books) AS book_id, f.user_ FROM facts f WHERE f.fact_date = '{DateTime.Now.ToString("dd.MM.yyyy")}') SELECT b.name AS book_name, ba.author_names AS book_author, CONCAT(u.surname, ' ', u.name, ' ', u.patronymic) AS who_return FROM fact_books fb JOIN books b ON fb.book_id = b.id JOIN book_authors ba ON fb.book_id = ba.book_id JOIN users u ON fb.user_ = u.id;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень возвращённых книг");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Книга");
                row.CreateCell(1).SetCellValue("Авторы");
                row.CreateCell(2).SetCellValue("Кем возвращена");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetString(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень_возвращённых_книг_за_{DateTime.Now.ToString("dd-MM-yyyy")}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetMostReadGenresListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "")
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("WITH current_month AS (SELECT date_trunc('month', CURRENT_DATE) AS start_of_month, date_trunc('month', CURRENT_DATE + INTERVAL '1 month') AS end_of_month), book_takings AS (SELECT b.id, unnest(b.genres) AS genre_id FROM facts f JOIN books b ON f.books @> ARRAY[b.id] WHERE f.taking_date >= (SELECT start_of_month FROM current_month) AND f.return_date <= (SELECT end_of_month FROM current_month) ) SELECT ROW_NUMBER() OVER (ORDER BY count(*) DESC) AS place, g.name AS genre_name, COUNT(*) AS book_count FROM book_takings bt JOIN genres g ON bt.genre_id = g.id GROUP BY g.name ORDER BY book_count DESC;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень самых читаемых жанров");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Место");
                row.CreateCell(1).SetCellValue("Жанр");
                row.CreateCell(2).SetCellValue("Число книг, взятых за месяц");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetInt16(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetInt32(2));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень_самых_читаемых_жанров_за_{DateTime.Now.ToString("MMMM")}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetMostReadAuthorsListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "") {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("WITH current_month AS (SELECT generate_series(date_trunc('month', CURRENT_DATE), date_trunc('month', CURRENT_DATE + INTERVAL '1 month') - INTERVAL '1 day', '1 day')::DATE AS month_day) SELECT ROW_NUMBER() OVER (ORDER BY count(DISTINCT b.id) DESC) AS place, a.name || ' ' || a.surname || ' ' || a.patronymic AS author, COUNT(DISTINCT b.id) AS book_count FROM facts f JOIN current_month cm ON f.taking_date <= cm.month_day AND f.return_date >= cm.month_day JOIN books b ON f.books @> ARRAY[b.id] JOIN UNNEST(b.authors) ua(author_id) ON TRUE JOIN authors a ON ua.author_id = a.id GROUP BY a.id, a.name, a.surname, a.patronymic ORDER BY book_count DESC;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Перечень самых читаемых авторов");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("Место");
                row.CreateCell(1).SetCellValue("Автор");
                row.CreateCell(2).SetCellValue("Число книг, взятых за месяц");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetInt16(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetInt32(2));
                }

                using (var file = new FileStream($"{ReportsPath}/Перечень_самых_читаемых_авторов_за_{DateTime.Now.ToString("MMMM")}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetFullBooksListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "")
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("WITH book_genres AS (SELECT b.id AS book_id, array_agg(g.name) AS genre_name FROM books b INNER JOIN genres g ON b.genres && ARRAY[g.id] GROUP BY b.id), book_authors AS (SELECT b.id AS book_id, array_agg(a.surname || ' ' || a.name || ' ' || a.patronymic) AS author_fullnames FROM books b INNER JOIN authors a ON b.authors && ARRAY[a.id] GROUP BY b.id) SELECT b.id AS book_id, b.name AS book_name, CASE WHEN b.status = 0 THEN 'Свободна' ELSE 'На руках у ' || COALESCE(u.surname || ' ' || u.name || ' ' || u.patronymic, '') END AS book_status, b.register_date AS book_register_date, array_to_string(array_agg(DISTINCT bg.genre_name), ', ') AS genre_names, array_to_string(array_agg(DISTINCT ba.author_fullnames), ', ') AS author_fullnames FROM books b LEFT JOIN book_genres bg ON b.id = bg.book_id LEFT JOIN book_authors ba ON b.id = ba.book_id LEFT JOIN users u ON b.status = u.id GROUP BY b.id, b.name, b.status, u.surname, u.name, u.patronymic;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Полный перечень книг");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("ID");
                row.CreateCell(1).SetCellValue("Название");
                row.CreateCell(2).SetCellValue("Статус");
                row.CreateCell(3).SetCellValue("Дата поступления");
                row.CreateCell(4).SetCellValue("Жанры");
                row.CreateCell(5).SetCellValue("Авторы");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetInt16(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                    row.CreateCell(3).SetCellValue(reader.GetDateTime(3).ToString("dd.MM.yyyy"));
                    row.CreateCell(4).SetCellValue(reader.GetString(4));
                    row.CreateCell(5).SetCellValue(reader.GetString(5));
                }

                using (var file = new FileStream($"{ReportsPath}/Полный_перечень_книг.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetFullUsersListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "")
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("SELECT id, surname, name, patronymic, email, register_date FROM users WHERE id > 0;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Полный перечень пользователей");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("ID");
                row.CreateCell(1).SetCellValue("Фамилия");
                row.CreateCell(2).SetCellValue("Имя");
                row.CreateCell(3).SetCellValue("Отчество");
                row.CreateCell(4).SetCellValue("Почта");
                row.CreateCell(5).SetCellValue("Дата регистрации");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetInt16(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                    row.CreateCell(3).SetCellValue(reader.GetString(3));
                    row.CreateCell(4).SetCellValue(reader.GetString(4));
                    row.CreateCell(5).SetCellValue(reader.GetDateTime(5).ToString("dd.MM.yyyy"));
                }

                using (var file = new FileStream($"{ReportsPath}/Полный перечень_пользователей.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetFullAuthorsListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "")
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("SELECT id, surname, name, patronymic FROM authors;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Полный перечень авторов");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("ID");
                row.CreateCell(1).SetCellValue("Фамилия");
                row.CreateCell(2).SetCellValue("Имя");
                row.CreateCell(3).SetCellValue("Отчество");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetInt16(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                    row.CreateCell(2).SetCellValue(reader.GetString(2));
                    row.CreateCell(3).SetCellValue(reader.GetString(3));
                }

                using (var file = new FileStream($"{ReportsPath}/Полный перечень_авторов.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
        private void GetFullGenresListButton_Click(object sender, RoutedEventArgs e)
        {
            if (ReportsPath == null || ReportsPath == "")
            {
                MessageBox.Show("Путь для сохранения отчётов не установлен!", "Ошибка", MessageBoxButton.OKCancel, MessageBoxImage.Error);
                return;
            }

            connection.Open();
            using (var command = new NpgsqlCommand("SELECT id, name FROM genres;", connection))
            {
                NpgsqlDataReader reader = command.ExecuteReader();
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Полный перечень жанров");
                int rowIndex = 0;

                var row = sheet.CreateRow(rowIndex);
                row.CreateCell(0).SetCellValue("ID");
                row.CreateCell(1).SetCellValue("Название жанра");

                while (reader.Read())
                {
                    rowIndex++;
                    row = sheet.CreateRow(rowIndex);
                    row.CreateCell(0).SetCellValue(reader.GetInt16(0));
                    row.CreateCell(1).SetCellValue(reader.GetString(1));
                }

                using (var file = new FileStream($"{ReportsPath}/Полный перечень_жанров.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(file);
                }
                MessageBox.Show($"Отчёт был успешно сформирован и сохранён в папку: {ReportsPath}", "Успех", MessageBoxButton.OKCancel, MessageBoxImage.Information);
            }
            connection.Close();
        }
    }
}
