import json
import keyboard
import pyautogui

MUTABLE_KEYS = ("download_btn_position", "student_name_position")


def get_position_on_spacebar():
    while True:
        keyboard.wait("space")
        break
    return pyautogui.position()


def click_on_position(position: tuple[int, int]):
    keyboard.wait("space")
    pyautogui.click(position)


def update_key_indexes(key_indexes: list[int]):
    with open("canva_coordinates.json") as read_json_file:
        coordinates = json.loads(read_json_file.read())

    try:
        for key_index in key_indexes:
            print(f"\n[UPDATING] {MUTABLE_KEYS[key_index]}")
            print("Move to the desired position and press spacebar to record it")
            recorded_position = get_position_on_spacebar()

            print(f"Recorded Position: {recorded_position}")
            print("Press spacebar to let computer click on recorded position for trial")
            click_on_position(recorded_position)
            print("Do you want to change the recorded position?")
            change_pos_choice = input("Enter 1 for yes, anything else for no - ")
            if change_pos_choice == "1":
                update_key_indexes([key_index])
                continue

            # Updating new position in the list, to update file at the end
            coordinates[MUTABLE_KEYS[key_index]] = list(recorded_position)
    except IndexError:
        print("Invalid index")

    with open("canva_coordinates.json", "w") as write_json_file:
        write_json_file.write(json.dumps(coordinates, indent=4))


def main():
    for index, key in enumerate(MUTABLE_KEYS):
        print(f"{index} For {key}")
    print("Enter indexes of keys seperated by space")
    key_indexes = list(map(int, input("Input - ").split()))
    update_key_indexes(key_indexes)


if __name__ == "__main__":
    main()
