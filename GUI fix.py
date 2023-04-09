import tkinter as tk
import openpyxl
import tkinter.messagebox as messagebox
from tkinter import ttk

# Load the Excel workbook
workbook = openpyxl.load_workbook("Moviesbase.xlsx")
worksheet = workbook.active

def check_age():
 
    age = int(age_entry.get())
    if age < 13:
        check_label.pack()
        name_entry.pack_forget()
        name_label.pack_forget()
        search_button.pack_forget()
        T3_label.pack_forget()
        genre_label.pack_forget()
        genre_combo.pack_forget()
        go_button.pack_forget()
        lookmovie_button.pack_forget()
        description_label.destroy()
        
 
    else:
        check_label.pack_forget()
        T3_label.pack()
        T3_label.place(x=0, y=0)
        genre_label.pack()
        genre_combo.pack()
        go_button.pack()
        age_label.pack_forget()
        age_entry.pack_forget()
        check_button.pack_forget()
        name_entry.pack_forget()
        name_label.pack_forget()
        search_button.pack_forget()
        description_label.destroy()
        
        
        
def lookmovie():
    recommended_label.config(text="")
    genre_label.pack_forget()
    genre_combo.pack_forget()
    lookmovie_button.pack_forget()
    go_button.pack_forget()
    name_entry.pack()
    name_label.pack()
    search_button.pack()
    go_back_button.pack()

#Define a function to get the description of the name
def get_description():
    name = name_entry.get() # Get the name from the user input
    found = False
    for row in worksheet.iter_rows(min_row=2, values_only=True): # Iterate over rows starting from row 2
        if row[0].lower() == name.lower():
            genre, ratings, year = row[1], row[2], row[3] # Get the values from the row
            description_label.config(text=f"Genre: {genre}\nRatings: {ratings}\nYear: {year}") # Update the label with the description
            found = True
    if not found:
        description_label.config(text= "Movie not found\n But we are expanding our range of movies! Check back in awhile!")


def trending_movies():
    result_label.config(text=f"Here are this week's Top 3 trending movies!\n1) Avatar: The Way of Water\n2) Top Gun: Maverick\n 3)Black Panther: Wakanda Forever")
    recommended_label.config(text="")
    for child in movie_frame.winfo_children():
        child.destroy()
    T3_label.pack()
    T3_label.place(x=0, y=0)
    genre_label.pack_forget()
    genre_combo.pack_forget()
    go_button.pack_forget()
    lookmovie_button.pack_forget()
    name_entry.pack_forget()
    name_label.pack_forget()
    search_button.pack_forget()
    go_back_button.pack()
    for child in movie_display_frame.winfo_children():
        if child.winfo_class() == "Label":
            child.destroy()
    recommended_movies.clear()
    # destroy the widgets that are no longer needed
    for widget in [movie_label, description_label]:
        widget.destroy()

def show_movies():
    global recommended_movies
    global movie_label
    selected_genre = genre_combo.get()
    result_label.config(text=f"You have selected {selected_genre} genre.")
    if selected_genre == "Action":
        recommended_movies = ["Aquaman", "Thor", "Batman"]
        recommended_label.config(text=f"Here are some recommendations for you:")
        for movie in recommended_movies:
            movie_label = tk.Label(movie_display_frame, text=movie, font=("Arial", 12), fg="blue", cursor="hand2")
            movie_label.pack(side="left", padx=10)
            movie_label.bind("<Button-1>", lambda e, movie=movie: show_movie_description(movie))
        genre_label.pack_forget()
        genre_combo.pack_forget()
        go_button.pack_forget()
        lookmovie_button.pack_forget()
        name_entry.pack_forget()
        name_label.pack_forget()
        search_button.pack_forget()
        go_back_button.pack(side="bottom", pady=10)

    elif selected_genre == "Comedy":
        recommended_movies = ["Minions: The Rise of Gru", "Deadpool", "Mr. Bean's Holiday"]
        recommended_label.config(text=f"Here are some recommendations for you:")
        for movie in recommended_movies:
            movie_label = tk.Label(movie_display_frame, text=movie, font=("Arial", 12), fg="blue", cursor="hand2")
            movie_label.pack(side="left", padx=10)
            movie_label.bind("<Button-1>", lambda e, movie=movie: show_movie_description(movie))
        genre_label.pack_forget()
        genre_combo.pack_forget()
        go_button.pack_forget()
        lookmovie_button.pack_forget()
        name_entry.pack_forget()
        name_label.pack_forget()
        search_button.pack_forget()
        go_back_button.pack(side="bottom", pady=10)

    elif selected_genre == "Sci-Fi":
        recommended_movies = ["Transformers: The Last Knight", "Avatar: The Way of Water", "Star Wars: The Rise of Skywalker"]
        recommended_label.config(text=f"Here are some recommendations for you:")
        for movie in recommended_movies:
            movie_label = tk.Label(movie_display_frame, text=movie, font=("Arial", 12), fg="blue", cursor="hand2")
            movie_label.pack(side="left", padx=10)
            movie_label.bind("<Button-1>", lambda e, movie=movie: show_movie_description(movie))
        genre_label.pack_forget()
        genre_combo.pack_forget()
        go_button.pack_forget()
        lookmovie_button.pack_forget()
        name_entry.pack_forget()
        name_label.pack_forget()
        search_button.pack_forget()
        go_back_button.pack(side="bottom", pady=10)

    elif selected_genre == "Animation":
        recommended_movies = ["Zootopia", "Shrek", "Turning Red"]
        recommended_label.config(text=f"Here are some recommendations for you:")

        for movie in recommended_movies:
            movie_label = tk.Label(movie_display_frame, text=movie, font=("Arial", 12), fg="blue", cursor="hand2")
            movie_label.pack(side="left", padx=10)
            movie_label.bind("<Button-1>", lambda e, movie=movie: show_movie_description(movie))

        genre_label.pack_forget()
        genre_combo.pack_forget()
        go_button.pack_forget()
        lookmovie_button.pack_forget()
        name_entry.pack_forget()
        name_label.pack_forget()
        search_button.pack_forget()
        go_back_button.pack(side="bottom", pady=10)

    elif selected_genre == "Documentary":
        recommended_movies = ["Fantastic Fungi", "Wanton Mee", "Puff: Wonders of the Reef"]
        recommended_label.config(text=f"Here are some recommendations for you:")

        for movie in recommended_movies:
            movie_label = tk.Label(movie_display_frame, text=movie, font=("Arial", 12), fg="blue", cursor="hand2")
            movie_label.pack(side="left", padx=10)
            movie_label.bind("<Button-1>", lambda e, movie=movie: show_movie_description(movie))
        genre_label.pack_forget()
        genre_combo.pack_forget()
        go_button.pack_forget()
        lookmovie_button.pack_forget()
        name_entry.pack_forget()
        name_label.pack_forget()
        search_button.pack_forget()
        go_back_button.pack(side="bottom", pady=10)

    else:
        recommended_label.config(text="")
        for child in movie_display_frame.winfo_children():
            child.destroy()  # remove the movie labels from the frame
        genre_label.pack()
        genre_combo.pack()
        go_button.pack()
        go_back_button.pack_forget()
    go_button.pack_forget()


movie_dict = {
    "Aquaman": "Aquaman is a 2018 American superhero film based on the DC Comics.\nThe film was directed by James Wan from a screenplay by David Leslie Johnson-McGoldrick and Will Beall.\nAquaman, who sets out to lead the underwater kingdom of Atlantis and stop his half-brother, King Orm, from uniting the seven underwater kingdoms to destroy the surface world.",
    "Thor": "Thor is a 2011 American superhero film based on the Marvel Comics.\nIt was directed by Kenneth Branagh.\nAfter reigniting a dormant war, Thor is banished from Asgard to Earth, stripped of his powers and his hammer Mjölnir. As his brother Loki (Hiddleston) plots to take the Asgardian throne, Thor must prove himself worthy.",
    "Batman": "Batman is a 2022 American superhero film based on the DC Comics.\n It was directed by Matt Reeves.\n Batman who has been fighting crime in Gotham City for 2 years.\n He has to uncover corruption while pursuing the Ridder.",
    "Minions: The Rise of Gru:": "Minions: The Rise of Gru is a 2022 American animated comedy film.\nIt is a sequel to the spin-off prequel Minions\nThe film was directed by Kyle Balda, co-direct by Brad Ableson and Jonathan del Val.\n11 years old Gru plans to become a super-villain with his Minions's help.\nThis let to a showdown with the Vicious 6.",
    "Deadpool": "Directed by Tim Miller.\n Wade Wilson, a former Special Forces operative who now works as a mercenary.\n His world came crashing down when evil scientist Ajax transform him into a Deadpool.",
    "Mr. Bean's Holiday": "A 2007 comedy film directed by Steve Bendelack.\n Mr Bean wins a trip to France.\n However, he is mistaken for both a kidnapper and an award-winning filmmaker.",
    "Transformers: The Last Knight": "It is a 2017 American science ficition action film based on Hasbro's Trnasformers toyline.\n5 years after the Hong Kong incident, Optimus Prime arrives on the ruins of Cybertron and meets Quintessa who brainwashes him into Nemesis Prime.\nHe was then send to the Earth to retrieve Merlin's staff, to restore the planet by taking Earth's energy core.",
    "Avatar:The Way of Water":"It is a 2022 American epic science fiction film.\n Directed by James Cameron\nFollows a blue-skinned humanoid Na'vi named Jake Sully as he and his family, under renewed human threat, seek refuge with the aquatic Metkayina clan of Pandora, a habitable exomoon on which they live.",
    "Star Wars: The Rise of Skywalker":"It is a 2019 American epic space opera film.\n Directed by J. J. Abrams.\nThe Rise of Skywalker follows Rey, Finn, and Poe Dameron as they lead the Resistance's final stand against Supreme Leader Kylo Ren and the First Order, who are aided by the return of the Galactic Emperor, Palpatine.",
    "Zootopia": "Zootopia is a 2016 American animated buddy cop comedy film produced by Walt Disney Animation Studios.\nThe film is about a rabbit police officer and a red fox con artist who form an unlikely partnership to uncover a conspiracy involving the disappearance of predator civilians within a mammalian metropolis.",
    "Shrek": "Shrek is a 2001 American animated comedy film based on the fairy tale picture book by William Steig.",
    "Turning Red": "Turning Red is a 2022 American animated coming-of-age comedy film directed by Domee Shi.",
    "Fantastic Fungi": "Fantastic Fungi is a 2019 documentary film directed by Louie Schwartzberg, which explores the world of fungi and their surprising benefits to our lives.",
    "Wanton Mee": "Wanton Mee is a 2021 Singaporean comedy film directed by Eric Khoo, which tells the story of a young woman who returns to Singapore to inherit her family's wanton noodle stall.",
    "Puff: Wonders of the Reef": "Puff: Wonders of the Reef is a 2021 nature documentary film directed by John Stoneman, which explores the fascinating world of coral reefs and their importance to our planet."
    # add more descriptions here
}

def get_movie_description(movie_name):
    description = movie_dict.get(movie_name)
    if description is None:
        return "Sorry, the movie description was not found."
    return f"\nIf interested, please use the following platforms: Netflix, Disney Plus\n{description}"

def show_movie_description(movie_name):
    for child in movie_frame.winfo_children():
        child.destroy()
    movie_label = tk.Label(movie_frame, text=movie_name, font=("Arial", 12), fg="blue")
    movie_label.pack(pady=5)
    description = get_movie_description(movie_name)
    description_label = tk.Label(movie_frame, text=description)
    description_label.pack(pady=5)

def go_back():
    recommended_label.config(text="")
    result_label.config(text="")
    for child in movie_frame.winfo_children():
        child.destroy()
    genre_label.pack()
    genre_combo.pack()
    go_button.pack()
    go_back_button.pack_forget()
    lookmovie_button.pack()
    name_entry.pack_forget()
    name_label.pack_forget()
    search_button.pack_forget()
    for child in movie_display_frame.winfo_children():
        if child.winfo_class() == "Label":
            child.destroy()
    recommended_movies.clear()
    description_label.destroy()
    movie_label.destroy()


def exit_program():
    if messagebox.askokcancel("Exit", "Are you sure you want to exit?"):
        messagebox.showinfo("Exit", "Thank you for using The Movie Guru. Have a nice day!")
        window.destroy()

# Create a window
window = tk.Tk()
window.title("The Movie Guru")
window.geometry("700x500")

# Create a label and entry for age input
age_label = tk.Label(window, text="Welcome\nFor verification, please enter your age:")
age_label.pack()
age_entry = tk.Entry(window)
age_entry.pack()

# Create a button to check age
check_button = tk.Button(window, text="Check", command=check_age)
check_button.pack()
check_label = tk.Label(window, text="Sorry, there are no available movies right now; please try again later.")
# Create a button to show recommended movies
go_button = tk.Button(window, text="Let's Go", command=show_movies) 
#create a button to look up for movie
lookmovie_button = tk.Button(window, text="Click to search a movie", command=lookmovie)
#lookmovie_button.pack()


# Create the user input label and entry widget
name_label = tk.Label(window, text="Check a movie out:")
name_label.pack()
name_entry = tk.Entry(window)
name_entry.pack()

# Create the button to trigger the name search
search_button = tk.Button(window, text="Search", command=get_description)
search_button.pack()

# Create the label to display the name description
description_label = tk.Label(window, text="")
description_label.pack()

#create a button to display top 3 trending movies
T3_label = tk.Button(window, text="Top 3 trending movies", command=trending_movies)

# Create a label and combo box for genre selection
genre_label = tk.Label(window, text="Please select a movie genre:")
genre_combo = ttk.Combobox(window, values=["Action", "Comedy", "Sci-Fi", "Animation", "Documentary"])
genre_combo.current(0) # Set default value to "Action"



# Create a label for the result
result_label = tk.Label(window, text="")
result_label.pack()

# Create a label for the recommended movies
recommended_label = tk.Label(window, text="")
recommended_label.pack()

# Create a frame for movie information
movie_frame = tk.Frame(window)
movie_frame.pack()

# Create a frame for movie display
movie_display_frame = tk.Frame(window)
movie_display_frame.pack()

# Create a button to go back to genre selection
go_back_button = tk.Button(window, text="Go Back", command=go_back)

#create a button to exit the program
exit_button = tk.Button(window, text="Exit", command=exit_program)
exit_button.pack(side="bottom", pady=10)


# Run the window
window.mainloop()