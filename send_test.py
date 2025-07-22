from src.controllers import SenderController

def run():
    sender_controller = SenderController()
    sender_controller.build(test_mode=True)

if __name__ == "__main__":
    run()