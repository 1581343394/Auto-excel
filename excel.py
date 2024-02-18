from gui import Application
if __name__ == "__main__":
    app = Application()
    try:
        app.mainloop()
    except KeyboardInterrupt:
        print("Program terminated by user")
        app.destroy()